"""ping Anthropic API เบาๆ เพื่อเช็คว่าเครดิตกลับมาหรือยัง (ไม่แตะ ZE)"""
import os, sys
from anthropic import Anthropic
try:
    r = Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"]).messages.create(
        model=os.environ.get("CLAUDE_MODEL", "claude-sonnet-4-6"),
        max_tokens=5, messages=[{"role": "user", "content": "ping"}])
    print(f"OK — เครดิตใช้ได้ | model={r.model} | tokens in/out={r.usage.input_tokens}/{r.usage.output_tokens}")
except Exception as e:
    print(f"FAIL — {e}")
    sys.exit(1)
