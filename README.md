# PPT 换皮工作台（北大红卡片模板）

把一份 PPT 丢进去，自动生成统一的「北大红卡片风格」成品——文字原样搬运，版式逐页替换。

- 自动识别内容类型：封面 / 目录 / 章节页 / 要点卡片 / 表格 / 结尾页
- 栏检测：网格布局不会被打乱顺序
- 老版 `.ppt` 自动先转 `.pptx`（需 LibreOffice）
- 本地 Flask 网站，文件不出本机

## 快速开始

```bash
pip install -r requirements.txt
python app.py
```

浏览器打开 http://127.0.0.1:5177 ，拖入 PPT，等待，下载。

## 依赖

- Python 3.10+，`flask` `python-pptx`
- 可选：LibreOffice（仅 `.ppt` 老格式需要，自动探测路径）
- 中文字体：Noto Sans SC / 思源黑体（缺失时会退回系统 CJK 字体）

## 直接调用引擎（不开网站）

```python
import engine
engine.plan_and_build("输入.pptx", "输出.pptx")
```

## 项目结构

```
engine.py          排版引擎：解析原稿 → 分类 → 套版生成
skin/base.pptx     版式底座
app.py             Flask 网站（上传/转换/下载）
templates/         前端页面
```

## 说明

- 样式来源：`skin/base.pptx`（北大红卡片风格参照件），改样式先改它或 `engine.py` 里的常量
- 转换在后台线程执行，网站页轮询进度
- 仅支持单文件上传，`.pptx` / `.ppt`
