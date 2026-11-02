# 致谢与来源声明

## 主要学习对象：PPT Master

本项目在**架构范式**上系统学习了 [PPT Master](https://github.com/hugohe3/ppt-master)
（作者 Hugo He，MIT 许可），并通过 DSH 生态中的
[`dsh-ppt-master`](https://www.npmjs.com/package/dsh-ppt-master) 插件得以近距离研究其
完整实现。以下设计直接来自对它的研读：

| 学到的机制 | PPT Master 中的形态 | 本项目中的落地 |
|---|---|---|
| **agent preset 隔离** | `@deepseek-ai/dsh-agent-preset` 声明行 + `cordis.patch.yml` | 同构：`preset-word-master` |
| **工具集分层** | 32 个 `ppt_*` 工具，一个脚本一个工具 | 13 个 `word_*` 工具 |
| **`extra_args` 逃生口** | 未建模旗标原样透传 | 每个脚本类工具都支持 |
| **每工具带 `cwd`** | 不假设进程工作目录 | 同构 |
| **skill provider + `${SKILL_DIR}` 展开** | 在 `get()` 阶段替换为真实绝对路径 | 同构，零依赖实现 |
| **SKILL.md 强制纪律块** | `[!CAUTION]` 全局执行纪律 + ⛔ BLOCKING 硬停 | 同构，10 条纪律 |
| **硬停闸门** | 八项确认（BLOCKING） | 规格冻结 + 质检 P0 + 审阅交接 |
| **视觉闸门** | `ppt_render` + `ppt_contact_sheet`，必须逐页看图 | `word_render` + `word_visual` |
| **本地 HTML 审阅界面** | `confirm_ui`（选参数）与 `svg_editor`（实时编辑器） | `review_ui`：深色三栏、点选段落、chips + 自由文本 |
| **mtime 漂移检测** | 服务端按 mtime 缓存 + 前端 2 秒轮询 + reload banner | 同构（`/api/mtime` + 横幅） |
| **暂存编辑 + 撤销栈** | 编辑先在预览可见，点「Apply changes」才写盘 | `/api/stage` 队列 + 原子提交 |
| **文件锁 + 端口避让 + 空闲看护** | `server_common.py` 的 `claim_lock` / `find_free_port` | 同构 |
| **面板布局与配色** | `svg_editor` 深色三栏（200 / flex / 280）、`#1a1a2e` 系 | 同构 |
| **素材转换分层** | `source_to_md/` 每格式一个转换器 | `convert.py` 五格式 |
| **模板系统** | brands / layouts / decks 三类 + 索引 | 简化为**样式集合**（见下） |
| **自检脚本** | `verify.mjs` 复刻 `installBundle` 硬校验 | 同构，67 项 |

## 有意做出的差异

这些不是遗漏，是判断后的取舍：

| 方面 | PPT Master | 本项目 | 原因 |
|---|---|---|---|
| 模板含义 | **版式骨架**（占位符摆位） | **样式集合**（字体/字号/边距） | Word 内容流是线性的，没有"版式"可套；真实需求是"新文档长得跟规范文档一样" |
| 中间产物 | 逐页手写 SVG | Markdown 稿 | 文本类内容用 Markdown 更自然，且可直接被审阅界面编辑 |
| 偏好记忆 | 无（风格在确认环节当场问） | **每轮注入系统提示词** | 写作偏好是长期约束，不能靠模型"想起来" |
| 视觉判据 | 基于 SVG bbox 几何 | 基于渲染后 PDF 的量测 | Word 没有 bbox，只能从渲染结果上量 |
| 特效能力 | 配音、动画、水印 | **不做** | Word 文档无对应概念 |

## 其他参考

- [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) —— 宿主平台
- [python-docx](https://github.com/python-openxml/python-docx) —— DOCX 读写
- [mammoth](https://github.com/mwilliamson/python-mammoth) —— DOCX → HTML
- [markdownify](https://github.com/matthewwithanm/python-markdownify) —— HTML → Markdown
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) —— PDF 提取与渲染
- [LibreOffice](https://www.libreoffice.org/) —— DOCX → PDF 渲染引擎

## 许可

本项目 MIT。PPT Master 亦为 MIT（Copyright (c) 2025-2026 Hugo He），其思想与
架构范式在本项目中以独立实现的形式被借鉴，未复制其源代码。

若你是 PPT Master 作者并认为本项目的致谢方式需要调整，请开 issue 告知。
