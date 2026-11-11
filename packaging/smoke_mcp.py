"""MCPモードの起動確認（スモークテスト）
=======================================

ビルドした exe（または開発中の entry.py）を `--mcp` で起動し、
MCP の initialize → tools/list を stdio で送って、stamp_execute が返るかを確かめる。
PyInstaller で必要なモジュールが同梱漏れしていると、ここで失敗する。

  python packaging/smoke_mcp.py dist/ltl-stamp/ltl-stamp.exe
  python packaging/smoke_mcp.py python entry.py        # 開発中
"""
import json
import subprocess
import sys
import threading

EXPECTED_TOOLS = {"stamp_execute"}
TIMEOUT_SEC = 60


def main(cmd):
    proc = subprocess.Popen(
        cmd + ["--mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
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
                raise RuntimeError(f"MCPサーバが応答前に終了しました:\n{err}")
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
        init = recv(1)
        if "result" not in init:
            raise RuntimeError(f"initialize 失敗: {init}")
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        tools = {t["name"] for t in recv(2)["result"]["tools"]}
    finally:
        timer.cancel()
        proc.kill()
        proc.wait()

    if tools != EXPECTED_TOOLS:
        raise RuntimeError(f"ツール一覧が想定と異なります: {sorted(tools)}")
    print(f"OK: MCPサーバ応答 / tools = {sorted(tools)}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    try:
        main(sys.argv[1:])
    except Exception as e:  # noqa: BLE001
        print(f"NG: {e}", file=sys.stderr)
        sys.exit(1)
