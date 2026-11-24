"""office.hc_bridge: Office's `herdr` inside a Herdr desktop app terminal (docs/herdr-app.md).

A fake `hc` keeps terminals in a JSON file. It also plays the pane's shell: each `hc get` takes the oldest queued
line the way the real pane host does, so the tests see a line start, run and end without a real terminal."""
from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

import pytest

from office import hc_bridge

FAKE_HC = r'''#!{python}
import json, os, sys
from pathlib import Path

state_file = Path(os.environ["FAKE_HC_STATE"])
data = json.loads(state_file.read_text()) if state_file.exists() else {"n": 0, "terminals": {}, "calls": [], "env": []}
args = [a for a in sys.argv[1:] if a != "--json"]
data["calls"].append(args)
data["env"].append({k: os.environ.get(k) for k in ("HC_TERMINAL_ID", "HC_SCOPE")})


def save():
    state_file.write_text(json.dumps(data))


def fail(kind, message, code=3):
    save()
    print(json.dumps({"error": {"kind": kind, "message": message}}))
    sys.exit(code)


def host(term):
    """One step of the pane's shell: take the oldest queued line, or end the one that is running."""
    q = term["labels"].get("office.q")
    if not q:
        return
    queue = Path(os.environ["OFFICE_DATA_HOME"]) / "hc-panes" / q
    mode = os.environ.get("FAKE_HC_AGENT", "idle")
    running = queue / "running"
    if running.exists():
        if mode == "exits":
            running.unlink()
            term["status"] = None
        return
    cmds = sorted(queue.glob("*.cmd"))
    if not cmds:
        return
    line = cmds[0].read_text()
    term.setdefault("lines", []).append(line)
    if term["labels"].get("office.cmd") == cmds[0].name:  # the agent's line: it holds the shell
        running.write_text(cmds[0].name + "\n")
        term["status"] = None if mode in ("silent", "exits") else {"app": "claude-code", "state": mode}
    cmds[0].unlink()


def view(term):
    return {k: term.get(k) for k in ("address", "state", "labels", "status", "cwd", "parent", "argv", "foreground")}


def find(name):
    term = data["terminals"].get(name)
    if term is None:
        fail("not_found", 'no terminal "%s"' % name)
    return term


cmd, rest = args[0], args[1:]
if cmd == "run":
    opts, argv, i = {"labels": {}}, [], 0
    while i < len(rest):
        a = rest[i]
        if a == "--":
            argv = rest[i + 1:]
            break
        if a == "--label":
            k, _, v = rest[i + 1].partition("=")
            opts["labels"][k] = v
            i += 2
        elif a in ("--cwd", "--parent", "--name"):
            opts[a] = rest[i + 1]
            i += 2
        else:
            opts[a] = True
            i += 1
    if os.environ.get("FAKE_HC_RUN_FAILS"):
        fail("spawn_failed", "cannot start")
    data["n"] += 1
    address = "t%d" % data["n"]
    data["terminals"][address] = {"address": address, "state": "live", "labels": opts["labels"], "status": None,
                                  "cwd": opts.get("--cwd"), "parent": opts.get("--parent"), "argv": argv,
                                  "foreground": {"cwd": opts.get("--cwd")}, "flags": sorted(k for k in opts if k.startswith("--"))}
    out = view(data["terminals"][address])
elif cmd == "get":
    term = find(rest[0])
    if term["state"] == "live":
        host(term)
    out = view(term)
elif cmd == "ls":
    want = dict(r.split("=", 1) for r in rest[1::2]) if rest[:1] == ["--label"] else {}
    rows = [t for t in data["terminals"].values() if t["state"] == "live"
            and all(t["labels"].get(k) == v for k, v in want.items())]
    out = {"terminals": [{"address": t["address"], "state": "live", "status": (t["status"] or {}).get("state")} for t in rows]}
elif cmd == "label":
    term = find(rest[0])
    for pair in rest[1:]:
        k, _, v = pair.partition("=")
        term["labels"][k] = v
    out = view(term)
elif cmd == "kill":
    term = find(rest[0])
    term["state"] = "finished"
    out = view(term)
elif cmd == "rm":
    find(rest[0])
    del data["terminals"][rest[0]]
    out = {"removed": rest[0]}
elif cmd == "read":
    term = find(rest[0])
    save()
    sys.stdout.write("".join("row %d\n" % i for i in range(1, 11)))
    sys.exit(0)
elif cmd in ("send", "keys"):
    term = find(rest[0])
    term.setdefault(cmd, []).append(rest[1:])
    out = {"status": None, "turn": 0}
else:
    fail("usage", "unknown command " + cmd, 2)
save()
print(json.dumps(out))
'''


class World:
    def __init__(self, tmp_path: Path, monkeypatch):
        self.state = tmp_path / "hc-state.json"
        self.data_home = tmp_path / "data"
        bindir = tmp_path / "bin"
        bindir.mkdir()
        fake = bindir / "hc"
        fake.write_text(FAKE_HC.replace("{python}", sys.executable))
        fake.chmod(0o755)
        monkeypatch.setenv("PATH", f"{bindir}{os.pathsep}{os.environ['PATH']}")
        monkeypatch.setenv("FAKE_HC_STATE", str(self.state))
        monkeypatch.setenv("OFFICE_DATA_HOME", str(self.data_home))
        monkeypatch.setenv("HC_TERMINAL_ID", "t0")
        monkeypatch.setattr(hc_bridge, "_host_shell", lambda: "/bin/sh")
        monkeypatch.setattr(hc_bridge.time, "sleep", lambda s: None)
        self.monkeypatch = monkeypatch
        self._capsys = None

    def data(self) -> dict:
        return json.loads(self.state.read_text())

    def calls(self, verb: str) -> list:
        return [c for c in self.data()["calls"] if c[0] == verb]

    def terminal(self, address: str) -> dict:
        return self.data()["terminals"][address]

    def herdr(self, *args: str) -> tuple[int, dict | str]:
        """One herdr call through the bridge: (exit code, parsed JSON or raw text)."""
        self._capsys.readouterr()
        code = hc_bridge.main(list(args))
        out = self._capsys.readouterr().out
        try:
            return code, json.loads(out)
        except ValueError:
            return code, out

    def split(self, **kw) -> str:
        code, out = self.herdr("pane", "split", "--pane", kw.get("pane", "t0"), "--direction", "right",
                               "--cwd", kw.get("cwd", "/work"), "--no-focus")
        assert code == 0, out
        return out["result"]["pane"]["pane_id"]

    def queue(self, pane: str) -> Path:
        return self.data_home / "hc-panes" / self.terminal(pane)["labels"]["office.q"]


@pytest.fixture
def world(tmp_path, monkeypatch, capsys):
    w = World(tmp_path, monkeypatch)
    w._capsys = capsys
    # The orchestrator's own terminal: a program the app reports, not an Office pane.
    w.state.write_text(json.dumps({"n": 0, "calls": [], "env": [], "terminals": {"t0": {
        "address": "t0", "state": "live", "labels": {}, "status": {"app": "claude-code", "state": "done"},
        "cwd": "/home", "foreground": {"cwd": "/home"}}}}))
    return w


# --- activation ---------------------------------------------------------------------------------------------

def test_the_bridge_serves_an_hc_terminal_and_never_a_real_herdr_pane(monkeypatch):
    for key in ("HC_TERMINAL_ID", "HERDR_ENV", "HERDR_PANE_ID", hc_bridge.FLAG):
        monkeypatch.delenv(key, raising=False)
    assert not hc_bridge.activate() and hc_bridge.herdr() == ["herdr"]  # no hc terminal: plain herdr

    monkeypatch.setenv("HC_TERMINAL_ID", "t7")
    assert hc_bridge.activate()
    assert os.environ["HERDR_ENV"] == "1" and os.environ["HERDR_PANE_ID"] == "t7"
    assert hc_bridge.herdr()[-2:] == ["-m", "office.hc_bridge"]
    assert hc_bridge.activate()  # a child that inherited all three is still served

    monkeypatch.setenv("HERDR_PANE_ID", "w1:p2")  # a real Herdr pane started from an hc terminal
    assert not hc_bridge.activate() and hc_bridge.herdr() == ["herdr"]


def test_the_bridge_can_be_switched_off(monkeypatch):
    monkeypatch.delenv("HERDR_ENV", raising=False)
    monkeypatch.setenv("HC_TERMINAL_ID", "t7")
    monkeypatch.setenv(hc_bridge.FLAG, "0")
    assert not hc_bridge.activate() and "HERDR_ENV" not in os.environ and hc_bridge.herdr() == ["herdr"]


def test_without_a_shell_to_host_a_pane_herdr_reads_as_missing(world, monkeypatch):
    monkeypatch.delenv("HERDR_ENV", raising=False)
    assert hc_bridge.activate() and hc_bridge.which_herdr()
    monkeypatch.setattr(hc_bridge, "_host_shell", lambda: None)
    assert hc_bridge.which_herdr() is None  # dispatch.herdr_usable() is then False: launches run headless


def test_office_calls_herdr_only_through_the_bridge():
    """A herdr call that names the binary itself would miss the bridge in a desktop app terminal."""
    src = Path(hc_bridge.__file__).parent
    bare = [f"{p.relative_to(src)}:{n}" for p in sorted(src.rglob("*.py")) if p.name not in ("hc_bridge.py", "install.py")
            for n, line in enumerate(p.read_text(encoding="utf-8").splitlines(), 1)
            if 'subprocess.run(["herdr"' in line or 'which("herdr")' in line]
    assert bare == []  # install.py runs real Herdr's own `integration` installer


# --- panes --------------------------------------------------------------------------------------------------

def test_a_split_starts_the_pane_host_beside_its_parent(world):
    pane = world.split(cwd="/work/tree")
    run = world.calls("run")[0]
    assert {"--bg", "--no-notify", "--outlive-parent"} <= set(run)  # never types notes into the orchestrator
    term = world.terminal(pane)
    assert term["parent"] == "t0" and term["cwd"] == "/work/tree"
    assert term["labels"]["office.pane"] == "1" and re.fullmatch(r"[0-9a-f]{32}", term["labels"]["office.q"])
    shell, host, queue = term["argv"]
    assert shell == "/bin/sh" and Path(host).read_text(encoding="utf-8") == hc_bridge.PANE_HOST
    assert Path(queue) == world.queue(pane) and world.queue(pane).is_dir()
    code, out = world.herdr("pane", "get", pane)
    assert code == 0 and out["result"]["pane"] == {"pane_id": pane, "tab_id": "hc", "cwd": "/work/tree",
                                                   "foreground_cwd": "/work/tree", "agent": None, "agent_status": "unknown"}


def test_a_split_from_a_pane_that_is_gone_lands_beside_the_caller(world):
    assert world.terminal(world.split(pane="t99"))["parent"] == "t0"


def test_a_failed_split_leaves_no_queue_behind(world, monkeypatch):
    monkeypatch.setenv("FAKE_HC_RUN_FAILS", "1")
    code, out = world.herdr("pane", "split", "--pane", "t0", "--direction", "right", "--cwd", "/work", "--no-focus")
    assert code == 1 and "error" in out
    assert [p for p in (world.data_home / "hc-panes").iterdir() if p.is_dir()] == []


def test_pane_run_queues_lines_in_order_for_the_pane_shell(world):
    pane = world.split()
    lines = [". '/r/agent.env' && cd '/work' && touch '/r/shell-ready'", "echo second"]
    for line in lines:
        assert world.herdr("pane", "run", pane, line)[0] == 0
    # The fake shell takes one line per `hc get`: what it ran, then what still waits, is the order sent.
    waiting = [q.read_text(encoding="utf-8") for q in sorted(world.queue(pane).glob("*.cmd"))]
    assert world.terminal(pane).get("lines", []) + waiting == lines
    assert world.calls("send") == []  # a line is never pasted into the shell


def test_a_closed_pane_is_gone_for_office(world):
    pane = world.split()
    queue = world.queue(pane)
    assert world.herdr("pane", "close", pane)[0] == 0
    assert not queue.exists() and pane not in world.data()["terminals"]
    code, out = world.herdr("pane", "get", pane)
    assert code == 1 and out["error"]["code"] == "pane_not_found"  # dispatch._pane_exists greps this word


def test_pane_read_returns_plain_text_and_visible_rows_only(world):
    pane = world.split()
    code, text = world.herdr("pane", "read", pane, "--source", "visible", "--lines", "3")
    assert code == 0 and text == "row 8\nrow 9\nrow 10\n"
    assert world.calls("read")[-1] == ["read", pane, "--lines", "0"]
    world.herdr("pane", "read", pane, "--source", "recent-unwrapped", "--lines", "5000")
    assert world.calls("read")[-1] == ["read", pane, "--lines", "5000"]


def test_a_rename_is_kept_as_a_label(world):
    pane = world.split()
    assert world.herdr("pane", "rename", pane, "repo", "T1", "#12", "[M]")[0] == 0
    assert world.terminal(pane)["labels"]["office.label"] == "repo T1 #12 [M]"


# --- agents -------------------------------------------------------------------------------------------------

def _start(world, pane, name="office-d1", kind="claude", args=("--model", "sonnet", "--add-dir", "/a b")):
    return world.herdr("agent", "start", name, "--kind", kind, "--pane", pane, "--", *args)


def test_agent_start_runs_the_harness_in_the_pane_shell_and_waits_for_it(world):
    pane = world.split()
    code, out = _start(world, pane)
    assert code == 0, out
    assert out["result"]["agent"] == {"name": "office-d1", "pane_id": pane, "agent": "claude", "status": "idle",
                                      "agent_status": "idle"}
    assert world.terminal(pane)["lines"] == ["claude --model sonnet --add-dir '/a b'"]  # herdr supplies the executable
    assert world.terminal(pane)["labels"]["office.agent"] == "office-d1"
    assert world.herdr("pane", "get", pane)[1]["result"]["pane"]["agent"] == "claude"
    by_name, by_pane = world.herdr("agent", "get", "office-d1")[1], world.herdr("agent", "get", pane)[1]
    assert by_name == by_pane and by_name["result"]["agent"]["pane_id"] == pane  # Office uses either


def test_a_second_agent_is_refused_while_the_pane_holds_one(world):
    pane = world.split()
    assert _start(world, pane)[0] == 0
    code, out = _start(world, pane, name="office-d2")
    assert code == 1 and out["error"]["code"] == "agent_pane_busy"  # Office then splits a fresh pane


@pytest.mark.parametrize("mode", ["blocked", "silent"])
def test_a_claude_held_on_a_startup_screen_is_not_ready_and_reads_blocked(world, monkeypatch, mode):
    """Claude reports a status as soon as its interface is up. One that runs without a word is on a trust or
    consent dialog: Office holds the pane for an answer instead of typing the brief into the dialog."""
    monkeypatch.setenv("FAKE_HC_AGENT", mode)
    monkeypatch.setattr(hc_bridge, "_START_QUIET", 0)
    pane = world.split()
    code, out = _start(world, pane)
    assert code == 1 and out["error"]["code"] == "agent_not_ready"
    code, out = world.herdr("agent", "get", "office-d1")  # still registered: startup.hold polls it
    assert code == 0 and out["result"]["agent"]["status"] == "blocked"


def test_a_harness_that_reports_nothing_counts_as_started_once_it_runs(world, monkeypatch):
    monkeypatch.setenv("FAKE_HC_AGENT", "silent")
    monkeypatch.setattr(hc_bridge, "_START_QUIET", 0)
    pane = world.split()
    code, out = _start(world, pane, kind="codex", args=("--yolo",))
    assert code == 0 and out["result"]["agent"]["status"] == "unknown"
    assert world.terminal(pane)["lines"] == ["codex --yolo"]


def test_a_harness_that_exits_as_it_starts_fails_the_start(world, monkeypatch):
    monkeypatch.setenv("FAKE_HC_AGENT", "exits")
    pane = world.split()
    code, out = _start(world, pane)
    assert code == 1 and out["error"]["code"] == "agent_start_failed"
    code, out = world.herdr("agent", "get", "office-d1")
    assert code == 1 and "not_found" in out["error"]["code"]  # rerun.agent_alive reads this as gone
    assert world.herdr("pane", "get", pane)[1]["result"]["pane"]["agent"] is None  # a shell again


def test_an_agent_that_never_existed_is_not_found(world):
    code, out = world.herdr("agent", "get", "office-nope")
    assert code == 1 and "not_found" in out["error"]["code"]


def test_agent_list_names_office_agents_and_the_orchestrator_terminal(world):
    """Root wake finds the orchestrator's own Claude by its pane and prompts it by name."""
    pane = world.split()
    world.split()  # a shell pane with no agent
    _start(world, pane)
    code, out = world.herdr("agent", "list")
    assert code == 0 and out["result"]["agents"] == [
        {"name": "t0", "pane_id": "t0", "agent": "claude", "status": "done", "agent_status": "done"},
        {"name": "office-d1", "pane_id": pane, "agent": "claude", "status": "idle", "agent_status": "idle"}]


def test_a_turn_that_ended_on_an_error_is_an_idle_agent(world, monkeypatch):
    monkeypatch.setenv("FAKE_HC_AGENT", "error")
    pane = world.split()
    assert _start(world, pane)[1]["result"]["agent"]["status"] == "idle"


def test_a_prompt_is_sent_as_plain_text_without_naming_this_terminal(world, monkeypatch):
    monkeypatch.setenv("HC_SCOPE", "scope")
    pane = world.split()
    _start(world, pane)
    pointer = "Read and carry out the brief at /r/brief.md exactly. When done, run: office submit"
    assert world.herdr("agent", "prompt", "office-d1", pointer)[0] == 0
    assert world.terminal(pane)["send"] == [[pointer]]
    sent = [e for c, e in zip(world.data()["calls"], world.data()["env"]) if c[0] == "send"]
    assert sent == [{"HC_TERMINAL_ID": None, "HC_SCOPE": None}]  # else hc wraps the text in a sender tag


def test_keys_use_hc_names_and_digits_are_pressed_one_by_one(world):
    pane = world.split()
    _start(world, pane)
    for keys in (["Enter"], ["esc"], ["down"], ["12"], ["ctrl+c"]):
        assert world.herdr("pane", "send-keys", pane, *keys)[0] == 0
    assert world.terminal(pane)["keys"] == [["enter"], ["esc"], ["down"], ["1", "2"], ["ctrl+c"]]
    code, out = world.herdr("pane", "send-keys", pane, "PageDown")
    assert code == 1 and out["error"]["code"] == "unknown_key"  # startup.answer checks the exit code


def test_line_clearing_keys_never_reach_the_pane_shell(world):
    """Office clears a shell's typed line before it retries a setup line. Here no line is typed, and ctrl+c
    would end the pane's shell."""
    pane = world.split()
    for key in ("ctrl+c", "ctrl+u"):
        assert world.herdr("pane", "send-keys", pane, key)[0] == 0
    assert "keys" not in world.terminal(pane)


def test_an_own_tab_is_one_pane_and_closes_with_it(world):
    code, out = world.herdr("tab", "create", "--label", "office-abc", "--cwd", "/work", "--no-focus")
    assert code == 0
    tab, pane = out["result"]["tab"]["tab_id"], out["result"]["root_pane"]["pane_id"]
    assert world.herdr("tab", "get", tab)[0] == 0
    assert world.herdr("tab", "close", tab)[0] == 0 and pane not in world.data()["terminals"]
    code, out = world.herdr("tab", "get", tab)
    assert code == 1 and "not_found" in out["error"]["code"]  # dispatch._tab_exists greps this word


def test_a_command_herdr_has_and_the_bridge_lacks_fails_plainly(world):
    code, out = world.herdr("workspace", "list")
    assert code == 1 and out["error"]["code"] == "unsupported"
