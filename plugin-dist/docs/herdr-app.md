# Panes in the Herdr desktop app

Office shows each dispatch as an agent in a pane beside the orchestrator when it runs inside Herdr. The Herdr
desktop app is a different program from the `herdr` terminal tool: its terminals belong to herdr-core and are
driven by `hc`. Office serves both. In a desktop app terminal it opens its panes there, with no setup.

## How it works

- **Detection.** A desktop app terminal sets `HC_TERMINAL_ID`. Any `office` command started there, outside real
  Herdr, turns the bridge on for itself and its children (`office.hc_bridge.activate`): it sets `HERDR_ENV=1`
  and `HERDR_PANE_ID` to the terminal's address (`t7`). Real Herdr always wins: a pane with its own
  `HERDR_PANE_ID` is left to `herdr`.
- **One `herdr`, two backends.** Office's pane code is unchanged. Every `herdr` call goes through
  `hc_bridge.herdr()`, which is `herdr` itself, or `python -m office.hc_bridge` when the bridge is on. The module
  answers the `herdr` commands Office uses (`pane split|get|run|read|close|rename|send-keys|send-text`,
  `agent start|get|read|list|prompt|send-keys`, `tab create|get|close`) in herdr's own JSON shapes.
- **A pane is an hc terminal.** `pane split` runs `hc run --bg --no-notify --outlive-parent` with the
  orchestrator's terminal as parent. The app shows a terminal beside its parent, so the pane appears as a split
  in the tab you are watching, and goes when Office closes it.
- **The pane's shell takes lines from a queue.** `hc run` does not pass the caller's environment, and text
  pasted into a shell arrives garbled. So the terminal runs a small shell loop (`PANE_HOST`) that reads command
  lines from a queue directory under the data home (`hc-panes/`) and runs them with `eval`. The setup line
  (`. agent.env && cd <worktree>`) therefore lasts for the agent started after it, and the agent owns the
  terminal while it runs. On Windows the shell is Git for Windows' bash; without it Office launches headless.
- **Talking to the agent.** The brief pointer, `office prompt` and answers go through `hc send` and `hc keys`.
  They are sent without naming the sending terminal, so the agent sees plain text.
- **Status.** `agent get` reports what the program told hc: `idle`, `working`, `done` or `blocked` (`error`
  reads as `idle`). Claude Code reports as soon as its interface is up, so a Claude that runs without having
  reported is on a startup screen and reads as `blocked`. A harness that reports nothing reads as `unknown`
  while it runs.

## What you will see

- A new split for each executor and reviewer, labelled with `office.label` (repository, task, role, route).
- A startup screen that needs your decision keeps its pane. Claude Code's bypass-permissions warning and a
  trust prompt for a directory Office did not create are yours to answer: answer in the pane, or with the
  `office answer` line `office status` prints. Office waits `OFFICE_STARTUP_PROMPT_WAIT` seconds (default 900),
  then runs that dispatch headless and says so.
- Read a pane with `hc read <pane>`. The pane id Office prints is the hc address.

## Limits

- `OFFICE_HC_BRIDGE=0` turns the bridge off (dispatches then run headless, as outside Herdr).
- The app decides the layout: new panes join the orchestrator's row. Office's split direction is not used.
- hc reports no shell pid or process group, so `agent.pgid` is not recorded for a pane agent. Office stops a
  pane agent by closing its pane.
- `pane send-text` has no equivalent without Enter; the bridge sends the text with Enter, and Office's own
  Enter after it lands on an empty composer.
- `herdr integration ...` (Hermes) stays real Herdr's own installer and is not bridged.
