# PaperFit Phase 2 使用说明

在项目根目录准备 Python 3.10 或以上的虚拟环境。运行依赖为 lxml、pypdf，测试另用 python-docx 和 ReportLab。安装时可访问包仓库，处理论文在本机完成。

```sh
python3 -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

只运行工具可安装 requirements.txt。保留整个 paperfit 目录即可作为 Agent Skill 使用，SKILL.md 与 scripts、references 保持同级。Python 集成入口为 scripts/phase2.py 的 process(source, plan, output)。

## 处理一篇论文

```sh
python scripts/paperfit.py extract 规范.pdf 规则草案.json
python scripts/paperfit.py audit 初稿.docx 结构检查.json
python scripts/paperfit.py run 初稿.docx 已核对规则.json 排版后.docx
python scripts/paperfit.py audit 排版后.docx 排版后检查.json
```

以上文件名为示例。extract 也接受含文字的 DOCX、UTF-8 TXT。PDF 保留页码，DOCX 保留正文块和表格行，TXT 保留行号。支持例如“正文：宋体小四，单倍行距，首行缩进2字符”“图表按章节编号”“公式全文连续编号”“参考文献顺序编号，悬挂缩进2字符”。扫描页、图像内容、冲突或无法解析的措辞进入 pending，须人工核对。模板视觉推断与 OCR 属于另行处理的输入。

先读结构检查的 paragraph_map。段落从 0 计数，目录中的每个编号、题注或短段落都可能与标题相似，需要结合上下文核对。chapter 映射指向章开始段落，可用于手打章号。审查同一输入的已有排版是否符合规则，可在 audit 命令后追加 `--rules 已核对规则.json`。稿件编辑或改变图题位置后重新检查索引。

以下规则只演示参数结构。将索引换为当前文件的位置。建议把 audit 的 document.sha256 抄入顶层 source_sha256，以便拒绝用错版本。

```json
{
  "styles": {"body": {"font": "宋体", "size": 12, "line": 1, "indent_chars": 2}},
  "paragraph_roles": {"1": "body"},
  "pending": [],
  "structure": {
    "chapters": {"0": 1, "6": 2},
    "captions": "chapter",
    "position": {"表": "before", "图": "after"},
    "equations": "chapter",
    "references": {"entries": [9], "number": true, "hanging_chars": 2},
    "sync_refs": true
  }
}
```

- captions、equations 分别支持 continuous 或 chapter。无需处理的模块省略。图与表分别计数。position 可选，仅移动紧邻且唯一匹配的题注。title_index 可选，按此前验收的做法隔离总标题与纯图片段落。以输入稿为基准确认全部索引。
- 图表与公式使用 Word SEQ 域。章节前缀取自已确认的章号，每章第一项重置序号。移动整章、更改章号后重新运行章映射。Word 原生新增题注依赖本机同名标签，连续编号已在用户环境实测。
- 公式处理针对正文中的独立可编辑 OMML，以及其末尾独立的手打括号编号。单栏单节版心内使用无边框三列表格，左右等宽，中央保留数学对象，右侧为编号。行内公式、复杂公式段落及其他区域进入待检查清单。
- 文献使用经核对的 entries 数组，按顺序添加或调整方括号序号，使用“参考文献条目”样式。hanging_chars 为悬挂缩进字符数，style 可含 font、size、line。检测文中 [1]、[1,3]、[1-3] 与条目编号的对应关系，提示缺失目标、未引用条目及疑似缺年份。已有原生列表或改号会改变既有引文含义时须核对。作者、题名、期刊和 DOI 均保留原样。这是基础排版，完整 GB/T 7714 著录审核需按用户给定版本另行处理。
- sync_refs 开启后，同步可唯一定位的普通“图2”“表3”“式（2）”及有明确书签范围的简单 REF 显示值，保留原生域结构。对 PAGEREF、NOTEREF 检查书签存在性，显示值交给 Word。复杂开关、断裂书签、歧义或无法准确分隔的引用进入报告或阻止重编号。
- v0.3.1 起，run 最后默认执行原生引用转换：连续方括号文献序号升级为自动列表，唯一的手打图表/文献引用升级为 REF 域。顶层 `native_references: false` 可明确关闭，`sync_refs` 与此开关职责不同。组合文献引用仍仅检查，不自动转换。详见 [原生交叉引用](native_references.md)。
- run 通过后生成新 DOCX 与同名 .report.json。status 为 needs_review 时查看 pending。文本账本逐段核对允许的编号与引用变动，另校验公式、原图、原表及包部件。对无法安全执行的请求，命令以状态码 2 退出并显示原因。输出和报告路径均须为新路径。

## 验收方法

运行全部测试可重复检查 PDF 真实文本抽取、Word 表格规范冲突、图表相邻匹配与位置、公式章号、文献条目、引用失效与同步、重复运行和原稿保护。在 Word 中打开结构化样例，可对照表1-1、两处公式编号 (1-1)、(2-1)，以及正文中的表号、公式号和文末 [1]。查看公式内容及编号在更新域后的行为，再通过样式窗格检查文献条目。WPS 与其他模板需要在对应环境复验。
