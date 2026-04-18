# Claude Design Prompt Workspace

这个仓库整理了当前这版设计相关素材与脚本，主要包括：

- `Claude-Design-Sys-Prompt_中文翻译(1).txt`：Claude 设计系统提示词中文翻译稿
- `build_defense_ppt.py`：基础版答辩 PPT 生成脚本
- `build_defense_ppt_school_template.py`：学校模板风格答辩 PPT 生成脚本
- `template_media/`：PPT 使用的图片素材
- `template_export/`、`defense_export/`、`defense_v2_export/`：导出预览图片

## 运行环境

- Python 3.11+
- `python-pptx`

安装依赖：

```bash
pip install python-pptx
```

运行示例：

```bash
python build_defense_ppt.py
python build_defense_ppt_school_template.py
```

## 说明

- 当前脚本中的部分输入/输出路径写死为本机路径，换电脑使用前建议先检查并修改。
- 仓库保留了导出图片，方便直接查看当前版本效果。
