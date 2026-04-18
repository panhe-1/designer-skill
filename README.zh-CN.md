# Designer Skill

[English README](README.md)

Designer Skill 是一个围绕设计类提示词资产、答辩演示文稿生成脚本和视觉预览导出文件整理的轻量工作区。

当前仓库主要围绕两条内容线展开：

- 保存 Claude 设计系统提示词的中文翻译草稿
- 使用 Python 和 `python-pptx` 生成论文答辩 PPT

## 仓库内容

- `Claude-Design-Sys-Prompt_中文翻译(1).txt`
  面向设计系统场景的 Claude 提示词中文翻译草稿。
- `build_defense_ppt.py`
  基础版论文答辩 PPT 生成脚本。
- `build_defense_ppt_school_template.py`
  学校模板增强版论文答辩 PPT 生成脚本。
- `template_media/`
  PPT 脚本使用的图片素材。
- `template_export/`
  模板风格演示文稿的导出预览图。
- `defense_export/`
  基础版答辩 PPT 的导出预览图。
- `defense_v2_export/`
  增强版答辩 PPT 的导出预览图。

## 运行要求

- Python 3.11 或更高版本
- `python-pptx`

安装依赖：

```bash
pip install python-pptx
```

## 快速开始

运行基础版 PPT 生成脚本：

```bash
python build_defense_ppt.py
```

运行学校模板增强版 PPT 生成脚本：

```bash
python build_defense_ppt_school_template.py
```

## 项目结构

这个仓库保持了比较轻量、以文件为中心的组织方式。

- 提示词材料以源文本形式保存，方便继续迭代。
- 演示文稿脚本通过结构化布局代码生成 `.pptx` 文件。
- 各类导出目录保留了幻灯片预览图，便于不打开 PowerPoint 也能快速查看效果。

## 说明

- 当前 Python 脚本中的部分输入输出路径仍然针对原始本地环境编写，换机器运行前建议先检查并调整。
- 仓库保留了导出的预览图片，方便直接在 GitHub 中查看当前视觉方向。
- 如果后续要把这个项目进一步整理成可复用的 skill，建议下一步把提示词资产、运行脚本和生成产物拆分到更清晰的子目录中。

## 当前状态

这个仓库目前更接近“持续整理中的工作快照”，还不是完全封装好的最终成品。
