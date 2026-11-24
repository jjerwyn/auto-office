"""Herdr panes in the Herdr desktop app (docs/herdr-app.md).

Office shows a dispatch as an agent in a pane beside the orchestrator by driving the `herdr` command line
(`pane split`, `pane run`, `agent start`, `agent get`, ...). The Herdr desktop app is a different program: its
terminals belong to herdr-core, driven by `hc` (`hc run`, `hc get`, `hc send`, ...), and it shows a terminal
started from another one as a split beside it.

This module is the `herdr` Office talks to when it runs in such a terminal. Office's pane code is unchanged:
every `herdr` call it makes goes to `main` here, which answers in herdr's own shapes from `hc`.

A herdr pane is a shell that Office types lines into. Text pasted into a shell through `hc send` arrives garbled
(bracketed-paste markers), and `hc run` does not hand the caller's environment to the program. So a pane here is
an hc terminal running `PANE_HOST`: a shell loop that takes its command lines from a queue directory. A line runs
in that same shell, so the dispatch environment and working directory it sets stay for the agent started after
it, and the agent owns the terminal while it runs. Text for the agent itself (the brief pointer, answers) goes
through `hc send`, which is made for that.

`activate` turns the bridge on for this process and its children: only inside an hc terminal, never inside real
Herdr, and only where a POSIX shell can host a pane (Git for Windows' bash on Windows).
"""
from __future__ import annotations

import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path

FLAG = "OFFICE_HC_BRIDGE"  # "1" once activated (children inherit it); "0" turns the bridge off
_IS_WINDOWS = sys.platform == "win32"
_HC_TIMEOUT = 25  # Office waits 30 s on a herdr call
_START_TIMEOUT = 100  # Office waits 120 s on `agent start`
_START_QUIET = 12  # how long a starting program may stay silent before the bridge answers for it
# Harness kinds whose program reports its status to hc as soon as its interface is up. One that runs without
# having reported is still on a startup screen (a trust or consent dialog): Herdr calls that `blocked`.
_REPORTS_STATUS = {"claude"}
# hc's status words are the program's own (idle, working, done, blocked, error). Herdr has no `error`: a turn that
# ended on an error is an agent waiting for input.
_STATUS = {"idle": "idle", "working": "working", "done": "done", "blocked": "blocked", "error": "idle"}
_APP_KIND = {"claude-code": "claude"}  # hc's name for a program -> the herdr kind Office compares against
_KEYS = {"enter": "enter", "esc": "esc", "escape": "esc", "up": "up", "down": "down", "left": "left",
         "right": "right", "tab": "tab", "space": "space"}

# The pane's shell. It runs each queued line with `eval`, so a line's `cd` and `export` last, and a program the
# line starts (the agent) has the terminal until it exits. `running` names the line in progress: that is how the
# bridge knows an agent is still up when the program reports no status of its own.
PANE_HOST = r"""q="$1"
trap : INT
printf '[office pane]\n'
while :; do
  for f in "$q"/*.cmd; do
    [ -e "$f" ] || continue
    line="$(cat "$f")"
    printf '%s\n' "${f##*/}" > "$q/running"
    rm -f "$f"
    printf '$ %s\n' "$line"
    eval "$line"
    rm -f "$q/running"
  done
  sleep 0.2
done
"""


# --- activation (the Office process) ------------------------------------------------------------------------

def _hc() -> str | None:
    return shutil.which("hc")


def _host_shell() -> str | None:
    """The POSIX shell that hosts a pane, or None when this machine has none Office can use."""
    if not _IS_WINDOWS:
        return shutil.which("bash") or shutil.which("sh")
    from office import portable
    root = portable._git_root()
    if root is None:
        return None
    wrapper = root / "bin" / "bash.exe"  # sets up Git Bash's own PATH before the pane's first line runs
    return str(wrapper if wrapper.is_file() else root / "usr" / "bin" / "bash.exe")


def available() -> bool:
    """Whether this process runs in an hc terminal that real Herdr does not already serve. Environment only: the
    hook fast path calls this on every tool call."""
    terminal = os.environ.get("HC_TERMINAL_ID")
    if not terminal or os.environ.get(FLAG) == "0":
        return False
    return os.environ.get("HERDR_ENV") != "1" or os.environ.get("HERDR_PANE_ID") == terminal


def activate() -> bool:
    """Make this process and its children see the hc terminal as their Herdr pane. Safe to call repeatedly."""
    if not available():
        if os.environ.get(FLAG) == "1":
            os.environ.pop(FLAG)  # inherited from a terminal this process is no longer in
        return False
    os.environ[FLAG] = "1"
    os.environ["HERDR_ENV"] = "1"
    os.environ["HERDR_PANE_ID"] = os.environ["HC_TERMINAL_ID"]
    return True


def active() -> bool:
    return os.environ.get(FLAG) == "1" and bool(os.environ.get("HC_TERMINAL_ID"))


def herdr() -> list[str]:
    """The argv prefix of a herdr call: `herdr`, or this module when the bridge serves this process."""
    if not active():
        return ["herdr"]
    return [sys.executable, *(["-X", "utf8"] if _IS_WINDOWS else []), "-m", "office.hc_bridge"]


def which_herdr() -> str | None:
    """`shutil.which("herdr")` for the gates that ask whether herdr can be used here. With the bridge: `hc`, and
    only where a shell can host a pane, so a machine without one launches headless as it does without Herdr."""
    if not active():
        return shutil.which("herdr")
    return _hc() if _host_shell() else None


# --- hc ---------------------------------------------------------------------------------------------------

class HcError(Exception):
    def __init__(self, kind: str, message: str):
        super().__init__(message)
        self.kind = kind


def _run_hc(args: list[str], *, timeout: float = _HC_TIMEOUT, sender: bool = True) -> subprocess.CompletedProcess:
    hc = _hc()
    if not hc:
        raise HcError("hc_missing", "hc is not on PATH")
    env = None
    if not sender:
        # hc wraps text one terminal sends another in a tag naming the sender. Office's prompts are the user's
        # run speaking, not this terminal's program: send them as plain text.
        env = {k: v for k, v in os.environ.items() if k not in ("HC_TERMINAL_ID", "HC_SCOPE")}
    try:
        return subprocess.run([hc, *args], capture_output=True, text=True, encoding="utf-8", errors="replace",
                              timeout=timeout, env=env, stdin=subprocess.DEVNULL)
    except subprocess.TimeoutExpired as exc:
        raise HcError("timeout", f"hc {args[0]} timed out") from exc
    except OSError as exc:
        raise HcError("hc_missing", str(exc)) from exc


def _hc_json(args: list[str], **kw) -> dict:
    proc = _run_hc([*args, "--json"], **kw)
    try:
        data = json.loads(proc.stdout or "{}")
    except ValueError:
        data = {}
    if not isinstance(data, dict):
        data = {}
    if proc.returncode != 0 or "error" in data:
        err = data.get("error") if isinstance(data.get("error"), dict) else {}
        if not err:
            try:
                err = (json.loads(proc.stderr or "{}") or {}).get("error") or {}
            except (ValueError, AttributeError):
                err = {}
        raise HcError(str(err.get("kind") or "hc_failed"), str(err.get("message") or proc.stderr or proc.stdout).strip()[:300])
    return data


def _terminal(target: str) -> dict | None:
    """A live terminal's record, or None when hc has none (finished ones are gone for Office)."""
    try:
        term = _hc_json(["get", target])
    except HcError as exc:
        if exc.kind == "not_found":
            return None
        raise
    return term if term.get("state") == "live" else None


def _queue(term: dict) -> Path | None:
    name = (term.get("labels") or {}).get("office.q")
    return _queue_root() / name if name and re.fullmatch(r"[0-9a-f]{32}", name) else None


def _queue_root() -> Path:
    from office import paths
    return paths.data_home() / "hc-panes"


def _agent_running(term: dict) -> bool:
    """Whether the line that started this pane's agent is still running in its shell."""
    labels = term.get("labels") or {}
    queue = _queue(term)
    if not queue or not labels.get("office.agent") or not labels.get("office.cmd"):
        return False
    cmd = labels["office.cmd"]
    try:
        if (queue / "running").read_text(encoding="utf-8").strip() == cmd:
            return True
    except OSError:
        pass
    return (queue / cmd).exists()  # queued, not yet taken by the shell


def _status(term: dict) -> str:
    state = (term.get("status") or {}).get("state")
    if state is None and (term.get("labels") or {}).get("office.kind") in _REPORTS_STATUS and _agent_running(term):
        return "blocked"
    return _STATUS.get(state, "unknown")


def _kind(term: dict) -> str | None:
    """The agent kind in this terminal: Office's own for a pane it started, else what the program reported."""
    labels = term.get("labels") or {}
    if labels.get("office.pane"):
        return labels.get("office.kind") if _agent_running(term) else None
    app = (term.get("status") or {}).get("app")
    return _APP_KIND.get(app, app) if app else None


def _agent_name(term: dict) -> str:
    return (term.get("labels") or {}).get("office.agent") or term["address"]


def _pane_view(term: dict) -> dict:
    fg = term.get("foreground") or {}
    return {"pane_id": term["address"], "tab_id": _tab_id(term), "cwd": term.get("cwd"),
            "foreground_cwd": fg.get("cwd") or term.get("cwd"), "agent": _kind(term), "agent_status": _status(term)}


def _agent_view(term: dict) -> dict:
    status = _status(term)
    return {"name": _agent_name(term), "pane_id": term["address"], "agent": _kind(term), "status": status,
            "agent_status": status}


def _tab_id(term: dict) -> str:
    return (term.get("labels") or {}).get("office.tab") or "hc"


def _find(target: str) -> dict | None:
    """The live terminal a herdr target names: an Office agent name, else an hc address or name."""
    if target.startswith("office-"):
        rows = _hc_json(["ls", "--label", f"office.agent={target}"]).get("terminals") or []
        for row in rows:
            term = _terminal(row["address"])
            if term:
                return term
        return None
    return _terminal(target)


# --- herdr's command line ---------------------------------------------------------------------------------

class Refused(Exception):
    """A herdr error: printed as herdr prints one, exit 1."""
    def __init__(self, code: str, message: str = ""):
        super().__init__(message or code)
        self.code = code


def _ok(result: dict) -> int:
    print(json.dumps({"result": result}))
    return 0


def _options(args: list[str], valued: tuple[str, ...], flags: tuple[str, ...] = ()) -> tuple[dict, list[str]]:
    """Split herdr arguments into {option: value} and positionals; everything after `--` is positional."""
    opts, rest, i = {}, [], 0
    while i < len(args):
        a = args[i]
        if a == "--":
            rest += args[i + 1:]
            break
        if a in valued and i + 1 < len(args):
            opts[a] = args[i + 1]
            i += 2
            continue
        if a in flags:
            opts[a] = True
        else:
            rest.append(a)
        i += 1
    return opts, rest


def _write_host() -> Path:
    root = _queue_root()
    root.mkdir(parents=True, exist_ok=True)
    host = root / "pane-host.sh"
    try:
        current = host.read_text(encoding="utf-8")
    except OSError:
        current = None
    if current != PANE_HOST:
        tmp = root / f".pane-host.{os.getpid()}"
        tmp.write_text(PANE_HOST, encoding="utf-8", newline="\n")
        os.replace(tmp, host)
    return host


def _new_pane(cwd: str | None, parent: str | None, tab: str | None = None) -> dict:
    shell = _host_shell()
    if not shell:
        raise Refused("no_shell", "no POSIX shell can host a pane on this machine")
    host = _write_host()
    name = uuid.uuid4().hex
    queue = _queue_root() / name
    queue.mkdir(parents=True)
    args = ["run", "--bg", "--no-notify", "--outlive-parent", "--label", "office.pane=1", "--label", f"office.q={name}"]
    if tab:
        args += ["--label", f"office.tab={tab}"]
    if cwd:
        args += ["--cwd", cwd]
    # The app shows a terminal beside its parent. A parent that is gone is no place: fall back to the caller's.
    if parent and _terminal(parent) is None:
        parent = None
    parent = parent or os.environ.get("HC_TERMINAL_ID")
    if parent:
        args += ["--parent", parent]
    try:
        term = _hc_json([*args, "--", shell, host.as_posix(), queue.as_posix()])
    except HcError:
        shutil.rmtree(queue, ignore_errors=True)
        raise
    return term


def _enqueue(term: dict, line: str) -> str:
    queue = _queue(term)
    if queue is None or not queue.is_dir():
        raise Refused("pane_not_shell", f"{term['address']} is not an Office pane")
    name = f"{time.time_ns():020d}.cmd"
    tmp = queue / f".{name}.tmp"
    tmp.write_text(line, encoding="utf-8", newline="\n")
    os.replace(tmp, queue / name)
    return name


def _need(target: str | None, code: str) -> dict:
    term = _find(target) if target else None
    if term is None:
        raise Refused(code, f"{target or 'target'} not found")
    return term


def _pane(args: list[str]) -> int:
    sub, args = args[0], args[1:]
    if sub == "split":
        opts, rest = _options(args, ("--pane", "--direction", "--cwd", "--ratio"), ("--no-focus", "--focus", "--current"))
        term = _new_pane(opts.get("--cwd"), opts.get("--pane") or (rest[0] if rest else None))
        return _ok({"pane": _pane_view(term)})
    if sub == "process-info":
        opts, rest = _options(args, ("--pane",))
        _need(opts.get("--pane") or (rest[0] if rest else None), "pane_not_found")
        return _ok({"process_info": {}})  # hc reports no shell pid or process group
    target, args = (args[0] if args else None), args[1:]
    term = _need(target, "pane_not_found")
    if sub == "get":
        return _ok({"pane": _pane_view(term)})
    if sub == "run":
        _enqueue(term, " ".join(args))
        return _ok({})
    if sub == "read":
        return _read(term, args)
    if sub == "close":
        return _close(term)
    if sub == "rename":
        _hc_json(["label", term["address"], f"office.label={' '.join(args)}"])
        return _ok({})
    if sub == "send-keys":
        return _keys(term, args)
    if sub == "send-text":
        # hc has no paste without Enter. Office presses Enter after this itself; a second Enter on an empty
        # composer does nothing.
        return _send(term, " ".join(args))
    if sub == "focus":
        return _ok({})
    raise Refused("unsupported", f"pane {sub}")


def _read(term: dict, args: list[str]) -> int:
    opts, _ = _options(args, ("--source", "--lines"))
    lines = int(opts["--lines"]) if str(opts.get("--lines") or "").isdigit() else None
    source = opts.get("--source") or "recent"
    history = 0 if source == "visible" else (lines if lines is not None else 2000)
    proc = _run_hc(["read", term["address"], "--lines", str(history)])
    if proc.returncode != 0:
        raise Refused("pane_not_found", (proc.stderr or proc.stdout).strip()[:200])
    text = proc.stdout
    if source == "visible" and lines is not None:
        text = "\n".join(text.splitlines()[-lines:]) + ("\n" if text.endswith("\n") else "")
    sys.stdout.write(text)
    return 0


def _close(term: dict) -> int:
    address = term["address"]
    for step in ("kill", "rm"):
        try:
            _hc_json([step, address])
        except HcError as exc:
            if exc.kind == "timeout":
                raise
    queue = _queue(term)
    if queue is not None:
        shutil.rmtree(queue, ignore_errors=True)
    return _ok({})


def _keys(term: dict, keys: list[str]) -> int:
    pressed = []
    for key in keys:
        low = key.lower()
        if low in _KEYS:
            pressed.append(_KEYS[low])
        elif re.fullmatch(r"(ctrl|alt)\+.", low) or re.fullmatch(r"f\d{1,2}", low):
            pressed.append(low)
        elif key.isdigit():
            pressed += list(key)  # herdr takes "12" as the keys 1 and 2
        elif len(key) == 1:
            pressed.append(key)
        else:
            raise Refused("unknown_key", key)
    if (term.get("labels") or {}).get("office.pane") and not _agent_running(term):
        # Office clears a shell's line with ctrl+c / ctrl+u before it retypes a setup line. This pane's shell
        # reads its lines from the queue: there is no typed line, and ctrl+c would only end the shell.
        pressed = [k for k in pressed if k not in ("ctrl+c", "ctrl+u")]
    if pressed:
        _hc_json(["keys", term["address"], *pressed])
    return _ok({})


def _send(term: dict, text: str) -> int:
    _hc_json(["send", term["address"], text], sender=False)
    return _ok({})


def _agent(args: list[str]) -> int:
    sub, args = args[0], args[1:]
    if sub == "list":
        return _ok({"agents": _agents()})
    if sub == "start":
        return _agent_start(args)
    target, args = (args[0] if args else None), args[1:]
    term = _find(target) if target else None
    if term is None or not _kind(term):
        raise Refused("agent_not_found", f"{target or 'agent'} not found")
    if sub == "get":
        return _ok({"agent": _agent_view(term)})
    if sub == "read":
        return _read(term, args)
    if sub == "prompt":
        return _send(term, " ".join(args))
    if sub == "send-keys":
        return _keys(term, args)
    raise Refused("unsupported", f"agent {sub}")


def _agents() -> list[dict]:
    """Every live terminal with an agent in it: Office's own panes, and programs that report a status to hc
    (the orchestrator's own terminal is one: Office wakes it by name)."""
    panes = {row["address"] for row in _hc_json(["ls", "--label", "office.pane=1"]).get("terminals") or []}
    out = []
    for row in _hc_json(["ls"]).get("terminals") or []:
        if not row.get("status") and row["address"] not in panes:
            continue  # a plain program: no agent to report
        term = _terminal(row["address"])
        if term and _kind(term):
            out.append(_agent_view(term))
    return out


def _agent_start(args: list[str]) -> int:
    opts, rest = _options(args, ("--kind", "--pane"))
    name, argv = (rest[0] if rest else None), rest[1:]
    kind, pane = opts.get("--kind"), opts.get("--pane")
    if not name or not kind or not pane or not re.fullmatch(r"[a-z][a-z0-9_-]{0,31}", name):
        raise Refused("invalid_agent_name", "usage: agent start <name> --kind <kind> --pane <pane> -- <args>")
    term = _need(pane, "pane_not_found")
    if not (term.get("labels") or {}).get("office.pane") or _agent_running(term) or (term.get("status") or {}).get("state"):
        raise Refused("agent_pane_busy", "not an available shell")
    cmd = _enqueue(term, shlex.join([kind, *argv]))
    _hc_json(["label", term["address"], f"office.agent={name}", f"office.kind={kind}", f"office.cmd={cmd}"])
    queue = _queue(term)
    deadline, started_at, taken = time.time() + _START_TIMEOUT, None, False
    while time.time() < deadline:
        time.sleep(0.5)
        term = _terminal(term["address"])
        if term is None:
            raise Refused("pane_not_found", f"{pane} closed while the agent started")
        if not taken:
            taken = not (queue / cmd).exists()
            started_at = time.time() if taken else None
            continue
        if not _agent_running(term):
            raise Refused("agent_start_failed", f"{kind} exited as it started")
        state = (term.get("status") or {}).get("state")
        if state == "blocked":
            raise Refused("agent_not_ready", "agent is blocked during startup")
        if state:
            return _ok({"agent": _agent_view(term)})
        if time.time() - started_at >= _START_QUIET:
            if kind in _REPORTS_STATUS:
                raise Refused("agent_not_ready", "agent is blocked during startup")
            return _ok({"agent": _agent_view(term)})  # this harness reports nothing: running is all hc can say
    raise Refused("agent_not_ready", "timeout: the agent did not come up")


def _tab(args: list[str]) -> int:
    sub, args = args[0], args[1:]
    if sub == "create":
        opts, _ = _options(args, ("--workspace", "--label", "--cwd"), ("--no-focus", "--focus"))
        tab = f"tab-{uuid.uuid4().hex[:12]}"
        term = _new_pane(opts.get("--cwd"), None, tab=tab)
        return _ok({"tab": {"tab_id": tab}, "root_pane": _pane_view(term)})
    tab = args[0] if args else ""
    rows = _hc_json(["ls", "--label", f"office.tab={tab}"]).get("terminals") or [] if tab.startswith("tab-") else []
    if sub == "get":
        if not rows:
            raise Refused("tab_not_found", tab)
        return _ok({"tab": {"tab_id": tab, "pane_count": len(rows)}})
    if sub == "close":
        for row in rows:
            term = _terminal(row["address"])
            if term:
                _close_quiet(term)
        return _ok({})
    raise Refused("unsupported", f"tab {sub}")


def _close_quiet(term: dict) -> None:
    import contextlib
    import io
    with contextlib.redirect_stdout(io.StringIO()):
        _close(term)


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    try:
        if len(args) >= 2 and args[0] == "pane":
            return _pane(args[1:])
        if len(args) >= 2 and args[0] == "agent":
            return _agent(args[1:])
        if len(args) >= 2 and args[0] == "tab":
            return _tab(args[1:])
        raise Refused("unsupported", " ".join(args[:2]) or "no command")
    except Refused as exc:
        print(json.dumps({"error": {"code": exc.code, "message": str(exc)}}))
        return 1
    except HcError as exc:
        code = "timeout" if exc.kind == "timeout" else f"hc_{exc.kind}"
        print(json.dumps({"error": {"code": code, "message": str(exc)}}))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
