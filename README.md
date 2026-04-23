# excel-data

基于 openpyxl 和 pandas 的 Excel 数据处理工具。

## 环境要求

- Python 3.12+
- [uv](https://docs.astral.sh/uv/) 包管理器

## 快速启动

```bash
# 安装依赖
uv sync

# 运行
uv run python main.py
```

## 项目结构

```
excel-data/
├── main.py          # 主入口
├── data/            # Excel 数据文件
├── pyproject.toml   # 项目配置与依赖
└── .venv/           # 虚拟环境（uv 自动创建）
```

## 依赖

- **openpyxl** — Excel 文件读写
- **pandas** — 数据分析与处理