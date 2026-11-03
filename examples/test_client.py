"""测试客户端：模拟 MCP 客户端，把完整对话跑一遍（也当用法示例）

Windows 上默认控制台是 GBK，直接 print 中文或替换字符会 UnicodeEncodeError，
所以这里先把输出流重配成 UTF-8，并给子进程设 PYTHONIOENCODING。
"""
import json, os, subprocess, sys, pathlib

for stream in (sys.stdout, sys.stderr):
    try:
        stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

SERVER = pathlib.Path(__file__).parent.parent / "server.py"
PY = sys.executable
OFFICE = r"D:\deepseek harness工作台\office办公软件"

cases = [
    {"jsonrpc": "2.0", "id": 1, "method": "initialize",
     "params": {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "test", "version": "1"}}},
    {"jsonrpc": "2.0", "method": "notifications/initialized"},
    {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
    {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
     "params": {"name": "doc_read", "arguments": {"path": OFFICE + r"\word\家庭用电分析报告.docx", "limit": 600}}},
    {"jsonrpc": "2.0", "id": 4, "method": "tools/call",
     "params": {"name": "sheet_read", "arguments": {"path": OFFICE + r"\excel\家庭用电数据分析.xlsx", "max_rows": 8}}},
    {"jsonrpc": "2.0", "id": 5, "method": "tools/call",
     "params": {"name": "slide_read", "arguments": {"path": OFFICE + r"\ppt\家庭用电情况分析.pptx", "limit": 700}}},
    {"jsonrpc": "2.0", "id": 6, "method": "tools/call", "params": {"name": "nope", "arguments": {}}},
    {"jsonrpc": "2.0", "id": 7, "method": "tools/call", "params": {"name": "doc_read", "arguments": {"path": "不存在.docx"}}},
]

payload = "\n".join(json.dumps(c, ensure_ascii=False) for c in cases) + "\n"
env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
proc = subprocess.run([PY, str(SERVER)], input=payload, capture_output=True, text=True,
                      encoding="utf-8", errors="replace", timeout=120, env=env)

print("==== stderr（日志，应只有启动信息）====")
print(proc.stderr.strip()[:300] or "（空）")

print("\n==== stdout 响应（逐条）====")
for line in proc.stdout.strip().splitlines():
    try:
        r = json.loads(line)
    except json.JSONDecodeError:
        print("  ⚠ 无法解析:", line[:120]); continue
    rid = r.get("id")
    if "error" in r:
        print(f"[id={rid}] ❌ error {r['error']['code']}: {r['error']['message']}")
        continue
    res = r.get("result", {})
    if rid == 1:
        print(f"[id=1] initialize -> {res['serverInfo']} 协议 {res['protocolVersion']}")
    elif rid == 2:
        print(f"[id=2] tools/list -> {len(res['tools'])} 个工具: {[t['name'] for t in res['tools']]}")
    else:
        text = res["content"][0]["text"] if res.get("content") else ""
        flag = "❌" if res.get("isError") else "✅"
        print(f"\n[id={rid}] {flag} 返回 {len(text)} 字符:")
        for l in text.splitlines()[:8]:
            print("     " + l[:110])