---
name: word-master
description: >
  Word 文档生产流水线：偏好核对 → 规格冻结闸门 → 逐节写正文 → 质检门（P0 阻断）
  → 审阅界面 → 渲染闸门 → 导出 DOCX/PDF。把素材做成排版规范、可交付的 Word。
  当用户说「写份报告」「做个文档」「生成 Word」「写论文」「改稿」「排个公文」
  时使用。
whenToUse: 需要产出或改写 .docx / 正式文档 / 论文 / 报告 / 公文时。
metadata:
  version: 1.0.0
  license: "MIT"
---

# Word Master

把「想写什么」变成一份能交付的 Word 文档。工具负责确定性操作，本文件负责顺序、
闸门和禁止事项。

> [!CAUTION]
> ## 🚨 全局执行纪律（MANDATORY）
>
> **这是一条严格串行的流水线。以下规则优先级最高，违反任何一条都算执行失败：**
>
> 1. **串行执行** — 步骤必须按顺序做，上一步的产物是下一步的输入。非 BLOCKING 的
>    相邻步骤在前提满足后可以连续做完，不必等用户说"继续"。
> 2. **BLOCKING = 硬停** — 标 ⛔ 的步骤必须完整停下等用户明确回复，**不得替用户拍板**。
> 3. **禁止跨阶段打包** — 规格阶段不写正文，正文阶段不排格式。
> 4. **进门先验前提** — 每步顶部的 🚧 GATE 必须先满足才能开始该步。
> 5. **不投机执行** — 不为后续步骤"提前准备"内容（例如在规格阶段先写几段正文试试）。
> 6. **偏好是约束不是建议** — `word_pref` 读到的每一条都会被质检门真的查。
> 7. **逐节写，禁止一次吐完整篇** — 一次吐全文会导致前后风格漂移，且返工代价最大。
> 8. **P0 未清零不得交付** — 不要用"基本完成，就剩几个小问题"糊过去。
> 9. **交付前必须真的看图** — `word_render` 出的 PNG 要逐页看过。只看退出码就宣布
>    完成是最常见的失职：字体回退、表格错位、标题孤行在退出码里完全看不出来。
> 10. **不手改 docx 的 XML** — 脚本已覆盖版式需求；手改会破坏三线表与样式继承，
>    下次重建即被覆盖。

> [!IMPORTANT]
> ## 🌐 语言与沟通
>
> - **回复语言**跟随用户输入与素材。用户明确要求（如"用英文写"）优先。
> - **文档内容**用用户的语言；**术语**按偏好要求白话化，首次出现给一句解释。
> - 汇报进度时只讲**结论与卡点**，不要复述你执行了哪些命令。

> [!IMPORTANT]
> ## 🔌 与通用写作习惯的关系
>
> - 本 skill 是**项目专属工作流**，不是通用写作助手。
> - 不要默认套用"总-分-总""每节三段""先写摘要再写正文"这类通用套路 —— 结构与
>   推进方式以**用户冻结的规格**和**已注入的偏好**为准。
> - 与通用写作习惯冲突时，遵循本 skill，除非用户明确说否则。

## 硬性加载顺序

**先路径后命令**：把本文件所在目录记为 `SKILL_DIR`，每次调用前展开
`${SKILL_DIR}`。不要 `cd`、不要依赖 CWD。路径拿不到就问用户，不要搜索或猜。

1. 读本文件。
2. 调 `word_env` 体检。缺 Python 或 python-docx 时把补齐命令交给用户；
   LibreOffice 缺失只影响渲染闸门，可继续但要说明。
3. 调 `word_pref`（action=read）确认当前偏好。用户提过新要求就先 set 落盘再开工。
4. 读 [`references/flow.md`](${SKILL_DIR}/references/flow.md) 了解每步判据与返工点。
5. 按流水线执行，**不跳步**。

---

## 流水线

```
⓪ word_project    建项目 + 收素材               —（新项目或素材陆续进来时）
⓪b word_convert    素材转 Markdown            —（用户给的不是 md 时）
① word_env        体检                    🚧 解释器与依赖就绪
② word_pref       偏好核对                 🚧 用户新要求已落盘
   word_template  套模板（可选）             —（用户有规范文档时）
③ word_spec       规格冻结        ⛔ BLOCKING  等用户确认章节结构
④ 写 draft.md     逐节写正文                 —（遵守已注入的偏好）
⑤ word_build      构建 docx                  —
⑥ word_gate       质检门          ⛔ BLOCKING  P0 清零前不得继续
⑦ word_review     审阅界面          ⛔ BLOCKING  等用户在浏览器里改完/提完意见
⑧ word_render     渲染闸门 + 视觉判据         —（必须逐页看图）
⑨ word_export     导出定稿                   —
```

### ⓪ 建项目与收素材（`word_project`）

**新项目先 init**，它建出 `sources/` 与 `converted/` 两个目录并写 `project.json`。
重复 init 是**幂等**的：只补缺的目录，不动已有的 spec/draft。

素材进来时用 `import-sources`：本地文件默认**复制**（不 move，用户的原始文件不该
因为我们收素材就消失），Markdown/文本直接进 `converted/`，文档类进 `sources/` 并
标记待转换。

### 随时可查：`word_project action=validate`

这是本流水线的**自检入口**，不是可有可无的附赠功能。它会：

1. **推断当前阶段**：`empty` / `draft` / `frozen` / `built` / `reviewed` / `delivered`
2. **给出下一步**：每个阶段都有明确的 `nextStep`
3. **抓隐性漏步**——这几类是最容易出事的：
   - `gate-p0`：质检还有 P0 未清就想交付
   - `stale-gate`：`draft.md` 比质检报告新（改了稿没重跑门禁，报告已过期）
   - `unconverted-sources`：素材收了但没转
   - `spec-not-frozen`：已有产物但规格未冻结（产物不对应任何已确认规格）
   - `no-visual-report` / `no-render`：声称交付但没做视觉验收

**跨会话续做时先跑 validate**，比重新勘察目录快得多。`action=info` 是一行看进度。

### ②b 套模板（`word_template`，可选）

**只在用户有现成规范文档时需要。** Word 的"模板"是**样式集合**（字体、字号、
行距、页边距、标题样式），不是 PPT 那种版式骨架——内容流是线性的，没有"版式"
可套。真实需求是：单位/学校有一份规范文档，新文档要长得一样。

```
word_template action=inspect  template=规范.docx     # 先看这份文档的版式画像
word_template action=save     template=规范.docx id=单位公文
word_template action=apply    id=单位公文 project=<项目>
```

⚠️ **apply 要在 `word_spec freeze` 之前**。freeze 会快照版式，之后 apply 的模板
不生效（要重新 freeze）。版式优先级：**模板 > 偏好 > 内置默认**。

### ③ 规格冻结（⛔ 硬停）

### ⓪ 素材转换（`word_convert`）

**只在用户给的输入不是 Markdown 时需要。** 支持的来源：PDF / Word / Excel /
PowerPoint / HTML / 网页 URL。

判据与坑：

| 来源 | 结果 | 注意 |
|---|---|---|
| PDF | 按字号推断标题层级 | 是启发式，封面/页眉可能误收；**转完先扫一遍标题** |
| Word (.docx) | 保留标题、列表、表格、图片 | 内嵌图片会落到 `<name>_files/` |
| Word (.doc) | **不支持** | 让用户另存为 .docx |
| Excel | 每 sheet 一节，输出管道表格 | 首行当表头 |
| PowerPoint | 每页一节 | 标题占位符当 `##` |
| 网页 | 提取正文，去掉导航/页脚 | 需要外网可达 |

转出来的 Markdown 是**素材，不是定稿**。它通常带着原文档的结构噪声（多余的
一级标题、把表格当正文、段落被硬切）。第 ③ 步的规格冻结要基于转换结果重新
规划章节，第 ④ 步重写而不是照搬。

### ① 体检（`word_env`）

判据：解释器可用 + `docx` / `lxml` 已装。缺转换依赖时只影响对应格式，不阻断。

`word_spec init` 建草稿，补全 `sections`，然后**停下来问用户这四项**：

| 必须问 | 为什么 |
|---|---|
| 章节结构 | 最常返工的一项。猜出来的结构和用户脑子里的一定不一样 |
| 目标读者 | 决定术语密度与解释深度 |
| 文类 | 论文/报告/公文/说明书的版式与语气差异很大 |
| 目标字数 | 决定每节展开到什么程度 |

用户确认后才 `freeze`。**冻结前写正文是明确禁止的。**

冻结会快照当时的偏好与版式 —— 之后用户改偏好，已完成的项目不受影响。

### ⑥ 质检门（⛔ 硬停）

构建后立刻跑。退出码说话，不靠感觉：

| 退出码 | 含义 | 动作 |
|---|---|---|
| 0 | 全过 | 进第 ⑦ 步 |
| 4 | 只有 P1 | 可交付，建议逐条处理 |
| 3 | 有 P0 | **不能交付**，改 draft.md 重建重跑 |
| 2 | 无法检查 | 先看文件路径与是否存在 |

P0：全角括号、禁用词命中、缺规格章节、标题缺失、标题跳级。
P1：节首句不合句式、方法章出现具体数值。

**改 draft.md 再重建，不要手改 docx。**

### ⑦ 审阅界面（⛔ 硬停）

`word_review action=open` 起本地页面并**把 URL 给用户**。用户在页面上：

- 直接改字（写回 `draft.md`）
- 段落右上角 💬 提意见（写进 `review-notes.json`）
- 底部写整篇级意见
- 点"应用修复"接受质检门的自动修复建议

**交接判据**：等用户说"好了/继续"，或 `word_review action=notes` 能看到意见之后，
先 `action=stop` 关掉服务器，再读 `review-notes.json` 与 `draft.md` 按意见改。

⚠️ **用完必须 stop**。占着端口的僵尸进程比没有界面更糟。

### ⑧ 渲染闸门 + 视觉判据

`word_render` 出 `render/pages/*.png`，然后 `word_visual` 从 PDF 上量排版问题：

| 判据 | 抓什么 |
|---|---|
| `line-spacing-tight` | 基线间距 / 字号 < 1.05（固定行距常见） |
| `text-out-of-margin` | 文字越过页面边界 |
| `text-overlap` | 两段文字渲染后重叠 |
| `orphan-heading` | 标题落在页底、正文被推到下页 |
| `font-size-drift` | 各页正文字号不一致（字体回退征兆） |

退出码 0/2/3/4，与 `word_gate` 同样的语义。

⚠️ **机器判据不能替代看图。** 表格观感、图注对齐、公式基线这些量不出来，
但一眼能看出。`word_visual` 通过只说明"没有可量化的破损"。

LibreOffice 缺失时 `word_render` 会跳过并说明 —— 此时只能说"结构质检通过，
未做视觉验收"，不能说"已通过视觉验收"。

---

## 工具速查

| 工具 | 什么时候用 |
|---|---|
| `word_project` | 建项目、收素材、查进度与漏步（**跨会话先跑 validate**） |
| `word_convert` | 素材不是 Markdown 时（PDF/Word/Excel/PPT/网页） |
| `word_env` | 开工第一步 |
| `word_pref` | 读/写偏好；用户提新要求时 |
| `word_template` | 套现成 .docx 的版式（inspect/save/list/apply） |
| `word_spec` | init 建草稿、freeze 冻结、read 复查 |
| `word_build` | 从 draft.md 构建 docx |
| `word_gate` | 每次构建后必跑，P0 阻断交付 |
| `word_review` | 起/停/查审阅界面，读意见 |
| `word_render` | 交付前必跑，出 PDF 与逐页 PNG |
| `word_visual` | 渲染后量排版（行距/越界/重叠/孤行） |
| `word_export` | 过闸后导出定稿（可带 PDF） |
| `word_read` | 改稿前了解现有文档结构 |
| `word_pref_section` | 核对当前生效的偏好原文 |

所有脚本类工具都支持 `extra_args` 透传未建模旗标，以及 `cwd` 指定工作目录。

## 项目目录约定

```
<project>/
├── project.json        项目清单（阶段/素材登记）
├── sources/            原始素材
├── converted/          素材转换出的 Markdown
├── spec.json           规格（冻结后只读）
├── draft.md            正文草稿（模型写，审阅界面也可改）
├── output.docx         构建产物
├── final.docx          交付定稿
├── gate-report.json    质检报告
├── visual-report.json  视觉判据报告
├── review-notes.json   ← 用户在审阅界面提的意见（AI 必读）
├── review-result.json  ← 交接记录（status/p0/p1/decisions）
└── render/
    ├── output.pdf
    └── pages/*.png
```

详细判据、常见返工点、偏好结构说明见 `references/flow.md`。
