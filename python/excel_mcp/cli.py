from __future__ import annotations

import argparse
import json
import os
import secrets
import shutil
import signal
import socket
import ssl
import subprocess
import sys
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import IO, Sequence

from dotenv import dotenv_values

PACKAGE_ROOT = Path(__file__).resolve().parents[2]
ADDIN_ROOT = PACKAGE_ROOT / "addin"
ENV_PATH = PACKAGE_ROOT / ".env"
CERT_DIR = Path.home() / ".office-addin-dev-certs"
CERT_PATH = CERT_DIR / "localhost.crt"
KEY_PATH = CERT_DIR / "localhost.key"
RELAY_PORT = 8000
ADDIN_PORT = 3002
READINESS_TIMEOUT_SECONDS = 15
READINESS_POLL_SECONDS = 0.5
CERT_RENEW_COMMAND = "cd addin && npx office-addin-dev-certs install --days 365"


class CliError(RuntimeError):
  """Raised for expected CLI failures that should print a concise message."""


def build_parser() -> argparse.ArgumentParser:
  parser = argparse.ArgumentParser(prog="excel-mcp")
  subparsers = parser.add_subparsers(dest="subcommand")

  setup_parser = subparsers.add_parser("setup", help="Install local prerequisites for the Excel MCP package")
  setup_parser.add_argument("--force", action="store_true", help="Regenerate .env secret and rerun setup steps")
  setup_parser.set_defaults(func=command_setup)

  start_parser = subparsers.add_parser("start", help="Start the relay and add-in dev server")
  start_parser.set_defaults(func=command_start)

  mcp_parser = subparsers.add_parser("mcp", help="Run the MCP stdio server")
  mcp_parser.set_defaults(func=command_mcp)

  return parser


def main(argv: Sequence[str] | None = None, *, default_subcommand: str | None = None) -> int:
  parser = build_parser()
  args_list = list(sys.argv[1:] if argv is None else argv)
  if not args_list and default_subcommand:
    args_list = [default_subcommand]

  if not args_list:
    parser.print_help()
    return 1

  args = parser.parse_args(args_list)
  if not hasattr(args, "func"):
    parser.print_help()
    return 1

  try:
    return int(args.func(args) or 0)
  except KeyboardInterrupt:
    print("Interrupted.", file=sys.stderr)
    return 1
  except CliError as exc:
    print(f"Error: {exc}", file=sys.stderr)
    return 1
  except Exception as exc:
    print(f"Error: {exc}", file=sys.stderr)
    return 1


def command_setup(args: argparse.Namespace) -> int:
  steps = [
    "Checking prerequisites",
    "Writing package .env",
    "Installing editable Python package",
    "Installing add-in dependencies",
    "Installing Office add-in dev certificates",
    "Verifying Office add-in dev certificates",
    "Registering the Excel add-in manifest",
    "Printing Claude Code MCP config",
    "Finalizing setup",
  ]

  _print_step(1, steps)
  _check_prerequisites()

  _print_step(2, steps)
  secret = _ensure_package_secret(force=bool(args.force))

  _print_step(3, steps)
  _run_checked(
    [sys.executable, "-m", "pip", "install", "-e", "./python"],
    cwd=PACKAGE_ROOT,
    step="Editable Python install",
  )

  _print_step(4, steps)
  _run_checked(["npm", "install"], cwd=ADDIN_ROOT, step="Add-in dependency install")

  _print_step(5, steps)
  _run_checked(
    ["npx", "office-addin-dev-certs", "install", "--days", "365"],
    cwd=ADDIN_ROOT,
    step="Certificate install",
  )

  _print_step(6, steps)
  _run_checked(
    ["npx", "office-addin-dev-certs", "verify"],
    cwd=ADDIN_ROOT,
    step="Certificate verification",
  )

  _print_step(7, steps)
  _run_checked(
    ["npx", "office-addin-dev-settings", "register", "manifest.xml"],
    cwd=ADDIN_ROOT,
    step="Manifest registration",
  )

  _print_step(8, steps)
  _print_mcp_config(secret)

  _print_step(9, steps)
  print("Setup complete. Run `python3 -m excel_mcp start`, then open Excel.")
  return 0


def command_start(_: argparse.Namespace) -> int:
  env_values = _load_package_env()
  _ensure_certificates()
  _ensure_port_available(RELAY_PORT)
  _ensure_port_available(ADDIN_PORT)

  child_env = os.environ.copy()
  child_env.update(env_values)

  relay_proc: subprocess.Popen[str] | None = None
  addin_proc: subprocess.Popen[str] | None = None
  stop_event = threading.Event()
  received_signal: dict[str, int | None] = {"value": None}
  previous_handlers = {
    signal.SIGINT: signal.getsignal(signal.SIGINT),
    signal.SIGTERM: signal.getsignal(signal.SIGTERM),
  }

  def _handle_signal(signum: int, _frame: object) -> None:
    received_signal["value"] = signum
    stop_event.set()

  signal.signal(signal.SIGINT, _handle_signal)
  signal.signal(signal.SIGTERM, _handle_signal)

  try:
    relay_proc = _start_process(
      "relay",
      [
        sys.executable,
        "-u",
        "-m",
        "uvicorn",
        "excel_mcp.relay:app",
        "--host",
        "0.0.0.0",
        "--port",
        str(RELAY_PORT),
        "--ssl-keyfile",
        str(KEY_PATH),
        "--ssl-certfile",
        str(CERT_PATH),
      ],
      cwd=PACKAGE_ROOT,
      env=child_env,
    )
    addin_proc = _start_process(
      "addin",
      ["npm", "run", "dev-server"],
      cwd=ADDIN_ROOT,
      env=child_env,
    )

    _wait_for_readiness(
      processes={"relay": relay_proc, "addin": addin_proc},
      ports={"relay": RELAY_PORT, "addin": ADDIN_PORT},
      stop_event=stop_event,
    )
    if stop_event.is_set():
      _stop_processes([relay_proc, addin_proc], received_signal["value"] or signal.SIGTERM)
      return 0
    print("Relay on https://localhost:8000, Add-in on https://localhost:3002. Open Excel.")

    while True:
      if stop_event.wait(READINESS_POLL_SECONDS):
        _stop_processes([relay_proc, addin_proc], received_signal["value"] or signal.SIGTERM)
        return 0

      relay_code = relay_proc.poll()
      addin_code = addin_proc.poll()
      if relay_code is not None:
        raise CliError(f"Relay exited unexpectedly with code {relay_code}.")
      if addin_code is not None:
        raise CliError(f"Add-in dev server exited unexpectedly with code {addin_code}.")
  finally:
    signal.signal(signal.SIGINT, previous_handlers[signal.SIGINT])
    signal.signal(signal.SIGTERM, previous_handlers[signal.SIGTERM])
    _stop_processes([relay_proc, addin_proc], signal.SIGTERM)


def command_mcp(_: argparse.Namespace) -> int:
  from .mcp_server import _prepare_stdio_instance, mcp

  _prepare_stdio_instance()
  mcp.run()
  return 0


def _print_step(index: int, steps: list[str]) -> None:
  print(f"[{index}/{len(steps)}] {steps[index - 1]}...")


def _check_prerequisites() -> None:
  if sys.version_info < (3, 10):
    raise CliError(f"python3 >= 3.10 is required. Found {sys.version.split()[0]}.")

  node_version = _command_version(["node", "--version"], name="node")
  if node_version[0] < 18:
    raise CliError(f"node >= 18 is required. Found {'.'.join(map(str, node_version))}.")

  _command_version(["npm", "--version"], name="npm")


def _command_version(command: Sequence[str], *, name: str) -> tuple[int, int, int]:
  if shutil.which(command[0]) is None:
    raise CliError(f"{name} is required but was not found in PATH.")

  try:
    result = subprocess.run(command, capture_output=True, text=True, check=False)
  except OSError as exc:
    raise CliError(f"Failed to run {name}: {exc}") from exc

  if result.returncode != 0:
    output = (result.stderr or result.stdout or "").strip()
    raise CliError(f"Failed to run {name}: {output or f'exit code {result.returncode}'}")

  raw = (result.stdout or result.stderr).strip().lstrip("v")
  parts = raw.split(".")
  if not parts or not parts[0].isdigit():
    raise CliError(f"Could not parse {name} version from: {raw}")

  major = int(parts[0])
  minor = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
  patch = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 0
  return major, minor, patch


def _ensure_package_secret(*, force: bool) -> str:
  env_exists = ENV_PATH.exists()
  current_values = dotenv_values(ENV_PATH) if ENV_PATH.exists() else {}
  current_secret = str(current_values.get("EXCEL_MCP_SECRET") or "").strip()
  if ENV_PATH.exists() and not force:
    if not current_secret:
      raise CliError(f"{ENV_PATH} exists but EXCEL_MCP_SECRET is missing. Re-run setup with --force.")
    print(f"Using existing EXCEL_MCP_SECRET from {ENV_PATH}.")
    return current_secret

  secret = secrets.token_urlsafe(32)
  _write_env_key(ENV_PATH, "EXCEL_MCP_SECRET", secret)
  action = "Regenerated" if env_exists and force else "Generated"
  print(f"{action} EXCEL_MCP_SECRET in {ENV_PATH}.")
  return secret


def _write_env_key(path: Path, key: str, value: str) -> None:
  lines = path.read_text().splitlines() if path.exists() else []
  prefix = f"{key}="
  replaced = False
  new_lines: list[str] = []
  for line in lines:
    if line.startswith(prefix):
      new_lines.append(f"{key}={value}")
      replaced = True
    else:
      new_lines.append(line)
  if not replaced:
    new_lines.append(f"{key}={value}")
  path.write_text("\n".join(new_lines).rstrip() + "\n")


def _run_checked(command: Sequence[str], *, cwd: Path, step: str) -> None:
  try:
    result = subprocess.run(command, cwd=cwd, check=False)
  except OSError as exc:
    raise CliError(f"{step} failed: {exc}") from exc

  if result.returncode != 0:
    raise CliError(f"{step} failed with exit code {result.returncode}.")


def _print_mcp_config(secret: str) -> None:
  config = {
    "mcpServers": {
      "excel-addin": {
        "type": "stdio",
        "command": str(Path(sys.executable).resolve()),
        "args": ["-m", "excel_mcp", "mcp"],
        "env": {
          "EXCEL_MCP_SECRET": secret,
          "EXCEL_MCP_BACKEND_URL": "https://localhost:8000/api/mcp/execute",
        },
      }
    }
  }
  print("Paste this into Claude Code MCP settings:")
  print(json.dumps(config, indent=2))


def _load_package_env() -> dict[str, str]:
  if not ENV_PATH.exists():
    raise CliError(f"Missing {ENV_PATH}. Run `python3 -m excel_mcp setup` first.")

  values = {
    key: value
    for key, value in dotenv_values(ENV_PATH).items()
    if isinstance(key, str) and isinstance(value, str)
  }
  secret = values.get("EXCEL_MCP_SECRET", "").strip()
  if not secret:
    raise CliError(f"{ENV_PATH} is missing EXCEL_MCP_SECRET. Run `python3 -m excel_mcp setup --force`.")
  values["EXCEL_MCP_SECRET"] = secret
  return values


def _ensure_certificates() -> None:
  if not CERT_PATH.exists() or not KEY_PATH.exists():
    raise CliError(
      "Office add-in dev certificates are missing. "
      f"Run `{CERT_RENEW_COMMAND}`."
    )

  try:
    decoded = ssl._ssl._test_decode_cert(str(CERT_PATH))
  except Exception as exc:
    raise CliError(f"Could not inspect {CERT_PATH}: {exc}") from exc

  not_after_raw = decoded.get("notAfter")
  if not isinstance(not_after_raw, str):
    raise CliError(f"Could not read certificate expiry from {CERT_PATH}.")

  expires_at = datetime.strptime(not_after_raw, "%b %d %H:%M:%S %Y %Z").replace(tzinfo=timezone.utc)
  if expires_at <= datetime.now(timezone.utc):
    raise CliError(
      f"Office add-in dev certificate at {CERT_PATH} is expired. "
      f"Run `{CERT_RENEW_COMMAND}`."
    )


def _ensure_port_available(port: int) -> None:
  owner = _port_owner(port)
  if owner is not None:
    raise CliError(f"Port {port} is already in use by {owner}.")


def _port_owner(port: int) -> str | None:
  try:
    result = subprocess.run(
      ["lsof", "-nP", f"-iTCP:{port}", "-sTCP:LISTEN", "-Fpcn"],
      capture_output=True,
      text=True,
      check=False,
    )
  except OSError as exc:
    raise CliError(f"Failed to inspect port {port} with lsof: {exc}") from exc

  if result.returncode not in (0, 1):
    output = (result.stderr or result.stdout or "").strip()
    raise CliError(f"Failed to inspect port {port}: {output or f'exit code {result.returncode}'}")

  pid = ""
  command = ""
  name = ""
  for line in result.stdout.splitlines():
    if line.startswith("p") and not pid:
      pid = line[1:]
    elif line.startswith("c") and not command:
      command = line[1:]
    elif line.startswith("n") and not name:
      name = line[1:]
    if pid and command and name:
      break

  if not pid:
    return None

  listener = name or f"TCP *:{port} (LISTEN)"
  if command:
    return f"{command} (PID {pid}) on {listener}"
  return f"PID {pid} on {listener}"


def _start_process(label: str, command: Sequence[str], *, cwd: Path, env: dict[str, str]) -> subprocess.Popen[str]:
  try:
    process = subprocess.Popen(
      command,
      cwd=cwd,
      env=env,
      stdin=subprocess.DEVNULL,
      stdout=subprocess.PIPE,
      stderr=subprocess.STDOUT,
      text=True,
      bufsize=1,
      start_new_session=True,
    )
  except OSError as exc:
    raise CliError(f"Failed to start {label}: {exc}") from exc

  assert process.stdout is not None
  thread = threading.Thread(target=_stream_output, args=(label, process.stdout), daemon=True)
  thread.start()
  return process


def _stream_output(label: str, pipe: IO[str]) -> None:
  try:
    for line in iter(pipe.readline, ""):
      if line.endswith("\n"):
        sys.stdout.write(f"[{label}] {line}")
      else:
        sys.stdout.write(f"[{label}] {line}\n")
      sys.stdout.flush()
  finally:
    pipe.close()


def _wait_for_readiness(
  *,
  processes: dict[str, subprocess.Popen[str]],
  ports: dict[str, int],
  stop_event: threading.Event,
) -> None:
  deadline = time.monotonic() + READINESS_TIMEOUT_SECONDS
  while time.monotonic() < deadline:
    if stop_event.is_set():
      return

    for label, process in processes.items():
      exit_code = process.poll()
      if exit_code is not None:
        raise CliError(f"{label.capitalize()} exited before becoming ready with code {exit_code}.")

    if all(_port_is_open(port) for port in ports.values()):
      return

    stop_event.wait(READINESS_POLL_SECONDS)

  pending = [label for label, port in ports.items() if not _port_is_open(port)]
  raise CliError(f"Timed out waiting for {', '.join(pending)} to bind their ports.")


def _port_is_open(port: int) -> bool:
  try:
    with socket.create_connection(("127.0.0.1", port), timeout=0.25):
      return True
  except OSError:
    return False


def _stop_processes(processes: Sequence[subprocess.Popen[str] | None], signum: int) -> None:
  alive = [process for process in processes if process is not None and process.poll() is None]
  if not alive:
    return

  for process in alive:
    try:
      os.killpg(process.pid, signum)
    except ProcessLookupError:
      pass

  deadline = time.monotonic() + 5
  while time.monotonic() < deadline:
    remaining = [process for process in alive if process.poll() is None]
    if not remaining:
      return
    time.sleep(0.1)

  for process in alive:
    if process.poll() is None:
      try:
        os.killpg(process.pid, signal.SIGKILL)
      except ProcessLookupError:
        pass


if __name__ == "__main__":
  raise SystemExit(main())
