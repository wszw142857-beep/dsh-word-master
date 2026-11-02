# dsh-word-master

[![dsh-plugin](https://img.shields.io/badge/dsh-plugin-DeepSeek%20Harness-4D6BFE)](https://github.com/topics/dsh-plugin)
[![license](https://img.shields.io/badge/license-MIT-green)](./LICENSE)
[![node](https://img.shields.io/badge/node-%3E%3D22-brightgreen)]()
[![python](https://img.shields.io/badge/python-3.10%2B-blue)]()

> Word Master for DeepSeek Harness —— 用 **agent preset** 形式把「素材 → 规范 Word 文档」
> 做成一条有闸门的流水线，并且**记住你的写作偏好**。

## 🙏 架构范式学习自 PPT Master

本项目在架构上系统学习了 **[PPT Master](https://github.com/hugohe3/ppt-master)**
（作者 Hugo He，MIT 许可）—— 通过 DSH 生态的
[`dsh-ppt-master`](https://www.npmjs.com/package/dsh-ppt-master) 插件近距离研读了它的
完整实现，包括：

- **agent preset 隔离**：只在选中预设的会话里生效，不占其他会话上下文
- **工具集分层**：一个脚本一个工具，`extra_args` 透传未建模旗标
- **skill provider + `${SKILL_DIR}` 展开**：文档里的命令行可直接复制执行
- **`[!CAUTION]` 强制纪律块 + ⛔ BLOCKING 硬停**：流程纪律而非流程自动机
- **视觉闸门**：渲染后必须逐页看图，不能只看退出码
- **本地 HTML 审阅界面**：`confirm_ui`（选参数）+ `svg_editor`（深色三栏实时编辑器）
- **mtime 漂移检测 + reload banner**：AI 改过文件时主动提示，而不是让你看旧内容
- **暂存编辑 + 撤销栈**：改完先在预览可见，点提交才一次性写盘
- **文件锁 / 端口避让 / 空闲看护**：本地服务器必须能自己退出

完整对照表、以及**有意做出的差异**（Word 的"模板"为什么是样式集合而不是版式骨架、
为什么中间产物是 Markdown 而不是 SVG 等）见 **[ACKNOWLEDGEMENTS.md](./ACKNOWLEDGEMENTS.md)**。

本项目为独立实现，未复制 PPT Master 的源代码。

## 它解决什么问题

| 没有它 | 有它 |
|---|---|
| 每次都要重新说"括号用半角""别写'显著'" | 偏好落盘，每轮自动注入，质检门真的查 |
| 模型自说自话宣布完成 | 质检门 P0 阻断，退出码说了算 |
| 结构检查过但字体回退、表格错位 | 渲染闸门 + 视觉判据，还要你真的看图 |
| LaTeX 公式原样显示成 `$$...$$` | 转成 Word 原生公式，可双击编辑 |
| 一次吐完整篇，前后风格漂移 | 逐节写，返工只影响一节 |
| 用户丢来一份 PDF 就没法接 | `word_convert` 直接读 PDF/Word/Excel/PPT/网页 |
| 套了单位模板却没用上 | `word_template` 提取版式，且 freeze 不覆盖它 |

## 安装

```bash
dsh plugin --profile desktop add D:/dsh/plugins/dsh-word-master
```

或者手工写进 profile 的 `dependencies`：

```json
{ "dsh-word-master": "link:D:/dsh/plugins/dsh-word-master" }
```

并在 `dsh.profile.bundles` 里追加 `"dsh-word-master"`。

> 改 profile 的 JSON 必须用 Node 写回，PowerShell 5.1 的 `Set-Content -Encoding UTF8`
> 会写 BOM，导致 Host 启动失败。

装完重启 DSH，新建会话时预设选择器里出现 **「Word Master」**。
**不选它的会话完全不受影响**，不占任何上下文。

## 六步流水线

```
① word_env      体检（Python / python-docx / lxml / LibreOffice）
② word_pref     偏好核对（read / set / memory_text）
③ word_spec     规格冻结 ← 硬停，等用户确认
④ 写 draft.md   逐节写，遵守已注入的偏好
⑤ word_gate     质检门 ← 硬停，P0 阻断交付
⑥ word_review   HTML 审阅界面 ← 硬停，等用户逐段设决定
⑦ word_render   渲染闸门 ← 必须逐页看图
⑧ word_export   导出 DOCX / PDF
```

## 审阅界面（`word_review`）

左右分栏的本地页面，视觉语言与 `dsh-ppt-master` 的确认页同一套（色板、圆角、
阴影、编号徽章、底部固定操作栏）。

```
┌──────────────────────────────┬───────────────────────┐
│ ① 正文（主）                 │ ② 质检发现            │
│   直接改字 → 写回 draft.md   │   定位 / 应用修复 /   │
│   每段右侧按钮设处理方式      │   处理                │
│                              ├───────────────────────┤
│                              │ ③ 已记录的决定        │
│                              ├───────────────────────┤
│                              │ ④ 全局补充            │
└──────────────────────────────┴───────────────────────┘
```

### 逐段设置 AI 改动

点正文任意段落右侧的 `＋`，弹出该段的处理浮层。两个通道并存：

- **结构化选择**（`catalogs.json` 的 12 项）：保持不动 / 重写 / 压缩 / 展开 /
  拆分 / 并入上段 / 调整位置 / 整段删除 / 改语言 / 改逻辑 / 补证据 / 移到方法章。
  可多选（例如"压缩 + 改语言"）。
- **自由文本**：那一段具体要怎么改，随便写。选 chips 和不选 chips 都能提交。

质检发现会自动标出推荐项（★）——这是把 ppt-master「AI 标注 recommended」的机制
移植过来，但用确定性映射而不是模型现场生成，同样的问题总推荐同样的处理方式。

整篇级需求分两块：**全局补充**（chips + 自由文本）与逐段决定分开存，AI 读的时候
不会混。

### 数据契约

```jsonc
// review-notes.json —— AI 必须读这个
{
  "schemaVersion": 2,
  "decisions": [
    { "line": 3, "quote": "原文前 120 字", "verdicts": ["shorten"], "note": "两个数据合并成一句" }
  ],
  "globalIntents": ["tighten", "conclusion_callback"],
  "globalNote": "结论回扣地下水暴露主线"
}
```

```jsonc
// review-result.json —— 交接判据
{ "status": "confirmed", "p0": 0, "p1": 1, "decisions": 4, "globalIntents": 2 }
```

**`verdicts` 为空但 `note` 非空是合法状态**（用户只写了一句话，没归类）。

### 服务器生命周期

- 文件锁 `.review-lock.json` 防同项目重复启动
- 端口被占自动往后找（5070 起，最多 50 个）
- 空闲超时默认 1800 秒自动退出
- `word_review action=stop` 优雅关闭 —— **用完必须 stop**

## 两个硬停点

**规格冻结**：章节结构错了正文全废，而结构只在用户脑子里。冻结前写正文是明确禁止的。

**质检 P0 + 审阅交接**：全角括号、禁用词、缺章节、标题缺失、标题跳级。
退出码 3 = 阻断，改到 0 或 4 才有资格进审阅。

## 工具

| 工具 | 作用 |
|---|---|
| `word_project` | 项目生命周期：init / import-sources / validate / info |
| `word_convert` | PDF / Word / Excel / PPT / 网页 → Markdown |
| `word_env` | 环境体检 |
| `word_pref` | 读写偏好；生成可写入记忆空间的文本 |
| `word_template` | 从现成 docx 提取版式并套用 |
| `word_spec` | init / read / freeze 规格 |
| `word_build` | Markdown → DOCX |
| `word_gate` | 质检门（退出码 0/2/3/4） |
| `word_review` | HTML 审阅界面（起/停/查/读决定） |
| `word_render` | DOCX → PDF → 逐页 PNG |
| `word_visual` | 从 PDF 量排版判据 |
| `word_export` | 定稿导出（可带 PDF） |
| `word_read` | 读现有 DOCX 结构，改稿前用 |
| `word_pref_section` | 查看当前生效的偏好原文 |

### `word_project validate`：随时知道做到哪一步

它推断当前阶段（`empty`→`draft`→`frozen`→`built`→`reviewed`→`delivered`），
给出下一步，并抓这几类**隐性漏步**：

| 检查 | 抓什么 |
|---|---|
| `gate-p0` | 质检还有 P0 未清就想交付 |
| `stale-gate` | `draft.md` 比质检报告新（改了稿没重跑门禁） |
| `unconverted-sources` | 素材收了但没转 |
| `spec-not-frozen` | 已有产物但规格未冻结 |
| `no-visual-report` / `no-render` | 声称交付但没做视觉验收 |

跨会话续做时先跑它，比重新勘察目录快得多。

## 偏好记忆怎么设计的

这是与 ppt-master 最大的不同。写作偏好是**长期约束**，不能靠模型"想起来"。

### 两层结构

1. **本地偏好库**（`$DSH_HOME/word-master/preferences.json`）是唯一事实来源。
   结构明确、可直接编辑、可版本化。
2. **每轮注入**：偏好被渲染成系统提示词小节，挂在**每一次请求**上
   （`ctx.systemPrompt.section()`），不依赖模型是否记得查。

### 与 mnemon 记忆空间的分工

记忆空间是按需检索的：模型得**想起来**调用召回才有内容。写作偏好恰恰不能靠自觉。
所以：

- 本地偏好库负责"每轮在场"；
- 记忆空间负责"跨机器、长期档案"——`word_pref` 的 `memory_text` 动作把偏好渲染成
  可直接交给 `mnemon_remember` 的一段文字，由你在说"记住"时写入。

两边都不重复造轮子，也不让关键约束悬在一次召回上。

### 偏好落盘的时机

用户说"以后都用小四""别写'显著'"这类话时，**必须**：

```
word_pref { action: "set", patch: { formatting: { bodySizePt: 10.5 } } }
```

只在对话里答应是没用的——下一轮注入的还是旧偏好。

## 自检

```bash
node verify.mjs
```

50 项检查，覆盖：插件契约、工具 schema、skill 载荷与占位符展开、Python 语法与依赖、
脚本 CLI 可用性、**JS 与 Python 两侧默认偏好一致性**、发布卫生（无本机路径/凭据/垃圾）。

## 结构

```
dsh-word-master/
├── package.json         # dsh.bundle 声明
├── cordis.patch.yml     # 声明 agent preset「Word Master」
├── lib/
│   ├── index.js         # 插件入口：工具 + skill provider + 偏好注入
│   ├── tools.js         # 9 个 word_* 工具定义
│   ├── runner.js        # 解释器探测与脚本执行
│   └── preferences.js   # 偏好库读写与渲染
├── skills/
│   ├── word-master/     # 上游完整流程 + references/flow.md
│   └── word-master-dsh/ # 精简工具集入口
└── verify.mjs           # 自检
```

## 实现笔记（踩过的坑）

这些都是实测踩出来的，写在这里免得重犯：

1. **`python3` 可能是商店别名**。Windows 上 `WindowsApps\python3.exe` 是
   AppInstaller 的占位重定向器，执行时**不跑脚本、直接返回 49**。只看 `which` 命中
   就采用会得到假故障。`runner.js` 对每个候选做真实探测（`-c "import sys;print(...)"`）。

2. **Python 侧偏好必须合并默认值**。早期版本文件不存在时返回空 `{}`，导致质检门的
   禁用词表为空——**静默什么都不查**，看起来"全部通过"。现在两侧都合并默认值，
   且 `verify.mjs` 会比对两侧是否一致。

3. **节首句正则里的 `.{2,}?` 是错的**。"为了判断风险，需要先…"这种最常见写法，
   逗号后直接接"需要"，早期正则要求中间至少 2 个字符，把合规句误判为不合规。

4. **h3 不是正文段落**。按二级标题切节时，三级标题若被收进 body，会变成"首句"，
   导致节首句检查误报。

5. **块级公式段必须单倍行距**。正文 1.5 倍行距会把带分式的高行盒再乘 1.5，
   视觉上公式上下各多一大块空白。

6. **标题默认是 Word 主题蓝**。python-docx 默认模板的 Heading 样式带 `#2F5496`，
   技术文档里像超链接。构建时直接改样式定义成黑色。

7. **中文字体要设 `w:eastAsia`**。只设 `font.name` 写的是 `w:rFonts/@ascii`，
   Word 会用主题字体回退，看起来"设置没生效"。

## License

MIT。
