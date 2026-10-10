# 白石山PPT

这个仓库包含一个可生成白石山演示文稿的 Python 脚本，以及本次演示的结构化文案和图片来源清单。

## 目录结构

- `generate_pptx.py`：生成 PPTX 的主脚本
- `requirements.txt`：依赖项
- `slides_content.md`：每页演示文稿正文与备注
- `image_sources.md`：图片来源和下载链接
- `baishishan_presentation.pptx`：生成后的 PPTX（本仓库中生成后会出现在这里）

## 生成方法

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python generate_pptx.py
```

生成后会在仓库根目录输出：

- `baishishan_presentation.pptx`
- `assets/` 目录（自动下载的图片）

## 功能说明

- 12页白石山主题演示文稿
- 中文标题和要点
- 图片自动下载与插入
- 每页演讲备注（可在脚本中扩展）
- 可继续修改内容并重新生成

## 参考方向

本汇报基于白石山的常见旅游、地貌、生态和旅行路线介绍方向，适合用于：

- 旅游宣传
- 旅行分享
- 校园科普
- 会议汇报

如果你希望我继续细化为更强营销风、学术风或实景旅行日记风格，可以直接告诉我调整方向。