# Phase 3 使用与验收

在 PaperFit 目录安装依赖后，通过 scripts/paperfit.py 调用。所有修改命令要求新的输出路径，旁边生成 JSON 报告。处理默认在本地完成。

## 已实现的入口

| 命令 | 用途 | 范围 |
| --- | --- | --- |
| toc | 生成 Word 原生目录 | 已识别的一至三级标题，单节单栏稿 |
| paginate | 目录和正文分节，正文从 1 编号 | 现有 PaperFit 目录、单节、无既有页脚 |
| chapter | 指定章节、指定标题层级的样式修改 | 正文中的唯一一级标题定位。支持现有多节稿 |
| template-info | 提取模板的样式与分节设置 | 输出样式名称、继承、字体字号、各节页面设置 |
| layout | 按明确映射迁移纸张尺寸、方向和四边页边距 | 指定输入稿第几节及模板第几节 |
| compare | 比较两个 DOCX 的正文块、段落样式、样式定义及包部件 | 独立生成 JSON 差异报告 |
| repair-toc | 修复目录链接丢失的目标 | 页码 PAGEREF 指向唯一有效标题书签，且标题文字吻合时修复。歧义保留并报告 |
| flow | 图表题注同页、一级标题另节和页眉页脚独立 | 交付前运行，具体对象范围及排版验收见 page_flow.md |

## 局部修改

先读取论文的一级标题原文。以下示例使用已验收的第二章标题。字号以磅为单位，四号为 14 磅。

```sh
python3 scripts/paperfit.py chapter input.docx output.docx --chapter "2.化学与能源的联系" --level 2 --font 黑体 --size 14
python3 scripts/paperfit.py compare input.docx output.docx difference.json
```

新样式使用章节标题、原样式名、字体和字号组成清晰名称，继承原样式的大纲与段落设置，进入快速样式库。只清理指定属性的直接覆盖。Word 保存可能重写 styleId，重复执行时会核对有效格式并复用现有样式。

标题文字与一级标题重复、嵌套标题、目标标题带字符样式或复杂对象、目标样式直接绑定多级列表、文档保护及修订等情况会停止处理并说明原因。只按特定样式名收录的目录需先核对。目录使用大纲级别收录时，局部样式仍保留其层级。

## 模板与多节

先提取模板的真实设置，再明确选择模板样式名和节序号。

```sh
python3 scripts/paperfit.py template-info school.docx template.json
python3 scripts/paperfit.py chapter input.docx styled.docx --chapter "2.化学与能源的联系" --level 2 --template school.docx --style "学校二级标题"
python3 scripts/paperfit.py layout styled.docx laid-out.docx --section 2 --template school.docx --template-section 2
```

模板样式迁移覆盖中西文字体、字号，以及模板明确给出的对齐、段间距、行距、缩进、与下段同页等支持属性。主题字体无法确定、学校规范互相冲突时需要用户核对。模板中封面内容、浮动对象、文献管理器域及页眉页脚关系仍需按实际模板单独处理。

模板首行缩进与目标的悬挂缩进相冲突，或反向冲突时，会停止输出并提示复核。

layout 按节定位，完整性检查核对指定页面属性以外的文档结构。节序号从 1 开始。报告记录迁移前后值和模板哈希，便于复核。

## 目录链接修复

```sh
python3 scripts/paperfit.py repair-toc input.docx repaired.docx
```

修复范围为带 PaperFit 标记的目录中，超链接目标丢失但页码书签完整的条目。已有的两个有效目标相互冲突、文字不吻合或书签不闭合时列入 pending。查看报告 status 后再决定是否交付。

## 验收与边界

运行自动测试。

```sh
python3 -m pip install -r requirements-dev.txt
python3 -m unittest discover -s tests -v
```

原生编辑验收包括在目标应用更新整个目录、保存重开、检查样式库及页码跳转。Word 与 WPS 各自的分页结果作为该应用内目录数字的依据。修改字体字号或页面设置后，目录缓存页码应由排版应用重新计算。

当前功能已有真实《素材2》回归和构造模板的自动测试。具体学校的复杂模板仍须以该模板实测验收。自动测试中的段落与对象完整性检查、应用内视觉核对、用户验收分别记录。

compare 的 body_block_text_equal 对比正文块的普通文字。图片、公式及其他对象的变更需要结合 body_changes 和 package_changes 检查。Word 保存可能规范化 XML，差异报告会如实列出这些结构变化。
