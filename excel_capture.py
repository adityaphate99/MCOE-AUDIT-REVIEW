"""Bound Excel COM calls in a disposable Python process, never the Excel process.

Excel can block inside Open, UsedRange, or Close while showing an Office dialog.
A timer in that same thread cannot interrupt COM. The parent therefore reads
progress while enforcing a deadline on this helper process. Index2 and the user's
Excel application are never terminated.
"""

import os
import json
import pickle
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def run_capture_command(command, output, progress, errors, timeout, status_callback=None):
    """Run one reader with a hard deadline, including cleanup of a stuck helper."""
    if timeout <= 0:
        raise ValueError("Excel capture timeout must be positive.")
    started = time.monotonic()
    last_status = "Starting Excel reader"
    next_heartbeat = started + 1
    with open(errors, "w", encoding="utf-8") as error_stream, open(progress, encoding="utf-8") as updates:
        process = subprocess.Popen(
            command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=error_stream,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
        try:
            while True:
                while True:
                    position = updates.tell()
                    line = updates.readline()
                    if not line:
                        break
                    if not line.endswith("\n"):
                        updates.seek(position)
                        break
                    last_status = line.rstrip("\r\n")
                    if status_callback:
                        status_callback(last_status)
                elapsed = time.monotonic() - started
                if process.poll() is not None:
                    break
                if elapsed >= timeout:
                    raise TimeoutError(
                        f"Excel timed out after {timeout:g}s. Last operation: {last_status}. "
                        "The Python reader was stopped; Excel and Index2 were not terminated. "
                        "Check Excel for sign-in, password, link, or compatibility dialogs."
                    )
                if time.monotonic() >= next_heartbeat:
                    if status_callback:
                        status_callback(f"{last_status} ({int(elapsed)}s / {timeout:g}s limit)")
                    next_heartbeat = time.monotonic() + 1
                time.sleep(min(0.05, timeout - elapsed))
            if process.returncode != 0:
                error_stream.flush()
                with open(errors, encoding="utf-8") as stream:
                    details = stream.read(4000).strip()
                raise RuntimeError(f"Excel reader failed (exit {process.returncode}): {details or 'no error details'}")
            if not Path(output).is_file():
                raise RuntimeError("Excel reader exited without captured workbook data.")
            with open(output, "rb") as stream:
                return pickle.load(stream)
        finally:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=2)
            else:
                process.wait()


def capture_family(url, timeout, status_callback=None, assemblies=()):
    """Return a value-only snapshot from our own trusted, local reader process."""
    with tempfile.TemporaryDirectory(prefix="mcoe-excel-") as folder:
        output = Path(folder) / "capture.pickle"
        progress = Path(folder) / "progress.log"
        errors = Path(folder) / "errors.log"
        progress.touch()
        command = [
            sys.executable, "-B", str(Path(__file__).resolve()),
            url, str(output), str(progress), json.dumps(list(assemblies)),
        ]
        return run_capture_command(command, output, progress, errors, timeout, status_callback)


def capture_worker(url, output, progress, assemblies_json="[]"):
    from io import BytesIO

    def report(message):
        with open(progress, "a", encoding="utf-8") as stream:
            stream.write(message + "\n")

    report("Connecting to Excel")
    import pythoncom
    from QA2 import ExcelComSession

    pythoncom.CoInitialize()
    session = None
    try:
        session = ExcelComSession(status_callback=report)
        name, sheets, snapshot = session._capture_live_family(url, json.loads(assemblies_json))
        try:
            report(f"Preparing cached data for {name}")
            buffer = BytesIO()
            snapshot.save(buffer)
            with open(output, "wb") as stream:
                pickle.dump((name, sheets, buffer.getvalue()), stream, protocol=pickle.HIGHEST_PROTOCOL)
        finally:
            snapshot.close()
    finally:
        session = None
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    capture_worker(*sys.argv[1:])
