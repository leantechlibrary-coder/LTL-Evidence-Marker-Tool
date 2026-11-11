"""MCPサーバを実プロセスで起動し、stdio で応答するかを確かめる。

`python mcp_server.py` を起動して initialize → tools/list を送り、stamp_execute が
返ることを確認する（import 単体テストでは見えない、起動経路と stdio の不具合を拾う）。
"""
import json
import os
import subprocess
import sys
import threading

import pytest

pytest.importorskip("mcp")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TIMEOUT_SEC = 60


def test_mcp_server_responds_over_stdio():
    proc = subprocess.Popen(
        [sys.executable, os.path.join(ROOT, "mcp_server.py")],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        cwd=ROOT,
    )

    def send(msg):
        proc.stdin.write((json.dumps(msg) + "\n").encode("utf-8"))
        proc.stdin.flush()

    def recv(want_id):
        # 通知（id なし）は読み飛ばし、目的の応答だけ返す
        while True:
            line = proc.stdout.readline()
            if not line:
                err = proc.stderr.read().decode("utf-8", "replace")
                pytest.fail(f"MCPサーバが応答前に終了しました:\n{err}")
            msg = json.loads(line)
            if msg.get("id") == want_id:
                return msg

    timer = threading.Timer(TIMEOUT_SEC, proc.kill)
    timer.start()
    try:
        send({
            "jsonrpc": "2.0", "id": 1, "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "smoke-test", "version": "0"},
            },
        })
        assert "result" in recv(1)
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        tools = {t["name"] for t in recv(2)["result"]["tools"]}
    finally:
        timer.cancel()
        proc.kill()
        proc.wait()

    assert tools == {"stamp_execute"}
