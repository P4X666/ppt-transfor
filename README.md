# ppt-transfor
ppt json ppt





```python
uv run python main.py parse <pptx>           # 解析单个 PPT 为 JSON
uv run python main.py parse-all              # 解析 input/ 下所有 PPT
uv run python main.py render <json>          # 从 JSON 渲染 PPT
uv run python main.py render-all             # 渲染 out/json/ 下所有 JSON
uv run python main.py compare <a.pptx> <b.pptx>  # 对比两个 PPT
uv run python main.py roundtrip <pptx>       # 完整往返：解析→渲染→对比
uv run python main.py roundtrip-all          # 批量往返测试
uv run pytest tests/                         # 运行测试
```

完善该ppt的README.md文件