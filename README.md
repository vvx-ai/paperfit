# PaperFit v0.3.2

PaperFit 是本地运行的论文 DOCX 格式化工具和 Agent Skill。输入论文与格式规则，输出新的可编辑 DOCX 和修改报告。处理引擎与 Agent 调度分离，不调用外部模型服务。

截至 2026-11-07，Phase 3 按已实现、已测试范围结项。用户反馈学校模板验收“基本没问题”。当前交付覆盖原生样式、编号、目录、分节、局部修改和明确审核的模板适配；不同学校的复杂模板仍需逐份验证。结项范围见 [v0.3 发布与验收记录](references/release_v0.3.md)。

v0.3.1 增加默认的手打引用转换与参考文献自动编号：`run` 和 `reviewed-template` 在输出前，将可唯一匹配的“图1”“表1”“[2]”转换为原生 REF 域，将连续 `[1]…[n]` 文献序号转换为 Word 原生列表。无法确定目标时保留并报告，不按位置猜测。使用方法、范围和插入条目测试见 [原生交叉引用](references/native_references.md)。

v0.3.2 默认将正文文献引用设为上标，方括号和数字一起上标，保留原生 REF 关联及更新后的格式。图表引用与文末文献列表编号保持正常基线。该步骤同时处理新建引用和可确认目标的既有引用。

## 安装和快速验证

需要 Python 3.10 或以上。已在 Python 3.12 验证。解压交付包，进入其中的 paperfit 目录。

```sh
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements-dev.txt
python scripts/paperfit.py --version
python scripts/paperfit.py --help
python -m unittest discover -s tests -v
python scripts/smoke_test.py smoke-output
```

Windows 激活命令为 `.venv\Scripts\activate`。只处理文档、不运行测试或合成样例时，安装 `requirements.txt` 即可。安装依赖需要包仓库访问，文档处理在本地完成。运行环境还需自行安装宋体等要求的字体，以及用于最终分页验收的 Word 或 WPS。

`requirements-tested.txt` 记录本次独立环境验证的精确依赖版本。需要复现该依赖组合时，用 `python -m pip install -r requirements-tested.txt` 替代范围依赖安装；该清单不包含操作系统、字体或应用版本。

`smoke-output` 必须是新目录。冒烟测试会生成合成文档，调用 audit → run → toc → paginate → flow，并检查原稿保护、图片与表格保留、重复执行和覆盖拒绝。它不使用任何用户论文，也不代替视觉验收。合成样例和修改报告可在该目录检查。

## 处理真实论文

```sh
python scripts/paperfit.py extract requirements.pdf draft-rules.json
python scripts/paperfit.py audit input.docx audit.json
# 根据规范、audit 的段落映射和原稿哈希，审核并保存 approved-rules.json。
python scripts/paperfit.py run input.docx approved-rules.json formatted.docx
python scripts/paperfit.py flow formatted.docx final.docx
```

不要直接把未审核的规则草案用于全文修改。不同命令的索引体系可能不同，尤其模板清单使用原始 body 块索引；以对应文档说明为准。存在 pending 的规则先解决歧义；无法安全处理的对象保留并报告。

需要首次生成目录与正文页码，且输入满足单节、单栏、无既有页脚等条件时，使用以下流程。

```sh
python scripts/paperfit.py toc formatted.docx toc.docx
python scripts/paperfit.py paginate toc.docx pages.docx
python scripts/paperfit.py flow pages.docx final.docx
```

已有目录用 Word/WPS 更新整个目录。不要重复插入目录，也不要将 `paginate` 套用于未经核对的多节模板。`reviewed-template` 已集成其审核范围内的目录、页码与同页处理，无需额外套用上述单节流水线。

## 命令与支持范围

所有命令均可通过 `python scripts/paperfit.py COMMAND --help` 查看参数。

| 命令 | 当前能力 | 主要边界 |
| --- | --- | --- |
| parse / format | 基础文字规则、原生标题正文样式、页面边距 | parse 只识别明确句式；复杂结构用 run |
| extract | PDF、DOCX、TXT 规范提取，保留来源位置 | 扫描件无 OCR；歧义、冲突进入 pending |
| audit / run | 结构、格式审查与图表公式编号、文献基础排版、可唯一定位的引用同步 | 原生独立公式、明确条目与章节映射；完整 GB/T 著录审核未实现 |
| toc / paginate | 原生目录，目录正文分节及 PAGE 页码 | 限定单节输入及已知目录结构 |
| flow | 图表题注同页、一级标题下一页分节、各节页眉页脚独立 | 相邻唯一对象、内嵌图片及短表；长表与浮动对象需单独处理 |
| chapter | 指定章节与层级的独立样式修改 | 唯一一级标题定位，有限字体/字号或模板样式迁移 |
| template-info / layout | 模板样式与页面信息、明确节映射 | 不自动搬运封面、复杂对象和任意页眉页脚 |
| compare / repair-toc | 修改差异、保守修复目录链接 | 不保证修复任意异常 DOCX |
| reviewed-template | 中文宋体/西文 TNR、三线表、摘要关键词与参考文献、目录分节综合适配 | 目前限定已审核的四节原稿结构，清单绑定输入哈希 |
| crossrefs | 手打图表/文献引用转 REF，连续方括号文献序号转原生列表 | 图表须有原生题注样式及 SEQ 域；歧义保留并报告 |

进一步说明见 [结构化规则](references/phase2.md)、[局部修改及模板](references/phase3.md)、[同页分节](references/page_flow.md)、[学校模板审核适配](references/reviewed_template.md)。用户在 Word 中插入题注后需要统一直接格式时，现有维护入口 `python scripts/normalize_captions.py input.docx normalized.docx` 仍可使用，之后重新执行 flow 和预览。

## 编辑与完整性约定

- 原稿保留，修改输出使用不存在的新路径。主要修改命令生成同名 `.report.json`。
- 使用原生样式与域；原文内容只在已授权的编号、引用同步或精确格式说明删除范围内变动，不补写文献元数据。
- 本项目已确认图表与题注同页、一级标题另页并优先分节、目录与正文页码分开、参考文献标题按一级标题处理。这些是当前项目约定，不是所有学校的默认标准。
- 新增图表和章节后重新执行对应处理。Word 的新增样式、题注标签与手动直接格式仍受本机模板和用户编辑影响，不能靠 DOCX 永久禁止产生差异。
- 更新整个目录后检查样式、页码、链接及打印预览，保存后再检查内容与图片。XML 校验通过不等于 Word/WPS 视觉验收。
- 本项目默认检查并转换手打引用及文献列表。只使用底层 `format` 时，之后运行 `crossrefs`；`run` 与 `reviewed-template` 已内置。用户明确只要求局部样式修改时保留局部范围，不自动扩大为全文引用转换。

## 作为 Skill 安装

将交付包内的 `paperfit` 文件夹放到你的技能目录，例如 `~/.codex/skills/paperfit`。保留 SKILL.md、VERSION、scripts 和 references 的相对位置，并为调用脚本的 Python 安装依赖。已有同名技能时先备份并比较版本，再替换。此次结项只提供安装包，未更改本机已安装技能。

## 发布包与历史记录

GitHub 源码提交范围和自动检查见 [GitHub 仓库说明](references/github.md)。仓库应以本 paperfit 目录为根；首次推送后核对 GitHub Actions 结果。当前尚未选定项目许可证，未添加 LICENSE 文件。

交付包包含源码、Skill、说明和合成测试，不包含 output 中的用户论文、审核清单、原始模板、应用预览、缓存或虚拟环境。依赖不随包分发，也未发布到公共仓库或 PyPI。

在项目根目录运行 `python scripts/build_release.py releases` 可重新打包并生成 SHA-256 校验文件。同一版本的已有包不会覆盖。

早期 README 已存入工作区 `references/history/README_before_v0.3.md`，仅作历史记录，不作为当前使用说明，也不进入发布包。各阶段工作日志保留在工作区。
