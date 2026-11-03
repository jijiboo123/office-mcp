#!/usr/bin/env python3
"""office-mcp：把办公文档接给 AI 的 MCP 服务器。

协议：MCP over stdio，每行一个 JSON-RPC 2.0 消息。
手写实现（没用官方 SDK），总共不到 200 行，方便看清协议到底长什么样。

支持的方法：
  initialize               握手，返回协议版本与能力
  notifications/initialized 客户端确认（通知，不需要回复）
  tools/list               列出可用工具
  tools/call               调用工具
  ping                     心跳

启动：
  python server.py            # 正常模式，读 stdin 写 stdout
  python server.py --selftest # 自检：列出工具并跑一个最小调用
"""

import json
import sys
import traceback

import tools

PROTOCOL_VERSION = "2024-11-05"
SERVER_INFO = {"name": "office-mcp", "version": "1.0.0"}


def _log(message):
    """日志一律走 stderr —— stdout 只能出 JSON，否则客户端会解析失败"""
    print(f"[office-mcp] {message}", file=sys.stderr)


def _result(msg_id, payload):
    return {"jsonrpc": "2.0", "id": msg_id, "result": payload}


def _error(msg_id, code, message):
    return {"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}}


def handle(request):
    """处理一条请求，返回响应 dict；返回 None 表示不用回复"""
    method = request.get("method")
    msg_id = request.get("id")
    params = request.get("params") or {}

    if method == "initialize":
        return _result(msg_id, {
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": SERVER_INFO,
        })

    if method in ("notifications/initialized", "initialized"):
        return None      # 通知没有 id，按协议不回复

    if method == "ping":
        return _result(msg_id, {})

    if method == "tools/list":
        return _result(msg_id, {
            "tools": [
                {"name": name, "description": spec["description"], "inputSchema": spec["schema"]}
                for name, spec in tools.TOOLS.items()
            ]
        })

    if method == "tools/call":
        name = params.get("name")
        arguments = params.get("arguments") or {}
        spec = tools.TOOLS.get(name)
        if not spec:
            return _error(msg_id, -32602, f"没有这个工具: {name}")
        try:
            outcome = spec["func"](**arguments)
        except TypeError as err:
            return _result(msg_id, {
                "content": [{"type": "text", "text": f"参数不对: {err}"}],
                "isError": True,
            })
        except Exception as err:  # noqa: BLE001 - 工具里任何异常都不能弄死服务器
            _log(f"工具 {name} 异常: {traceback.format_exc()}")
            return _result(msg_id, {
                "content": [{"type": "text", "text": f"工具执行失败: {err}"}],
                "isError": True,
            })

        if isinstance(outcome, dict) and outcome.get("ok"):
            return _result(msg_id, {"content": [{"type": "text", "text": outcome["text"]}], "isError": False})
        message = outcome.get("error") if isinstance(outcome, dict) else str(outcome)
        return _result(msg_id, {"content": [{"type": "text", "text": message}], "isError": True})

    return _error(msg_id, -32601, f"不支持的方法: {method}")

def serve():
    """逐行读 stdin，逐行写 stdout。这是 MCP stdio 传输的全部"""
    _log(f"启动，协议版本 {PROTOCOL_VERSION}，可用工具 {len(tools.TOOLS)} 个")
    for raw in sys.stdin:
        line = raw.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except json.JSONDecodeError as err:
            response = _error(None, -32700, f"JSON 解析失败: {err}")
        else:
            try:
                response = handle(request)
            except Exception as err:  # noqa: BLE001
                _log(traceback.format_exc())
                response = _error(request.get("id"), -32603, f"内部错误: {err}")
        if response is None:
            continue
        sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
        sys.stdout.flush()


def selftest():
    """自检：不依赖客户端，把工具列出来并跑一次真实调用"""
    print("工具清单:")
    for name, spec in tools.TOOLS.items():
        print(f"  {name:14} {spec['description'][:40]}…")
    outcome = tools.TOOLS["doc_read"]["func"](
        path=sys.argv[2] if len(sys.argv) > 2 else "missing.docx"
    )
    print("\ndoc_read 自检结果:", "成功" if outcome.get("ok") else outcome.get("error"))
    if outcome.get("ok"):
        print(outcome["text"][:400])
    return 0


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        raise SystemExit(selftest())
    try:
        serve()
    except KeyboardInterrupt:
        _log("收到中断，退出")