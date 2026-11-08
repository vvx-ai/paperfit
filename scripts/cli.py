"""One discoverable CLI for the existing engines; legacy script APIs remain valid."""
import argparse
import json
from pathlib import Path
import sys

COMMANDS = {
    'parse': '解析明确句式的文字规范',
    'format': '基础段落与页面格式',
    'extract': '提取 PDF、DOCX、TXT 规范草案',
    'audit': '只读结构、引用和格式审查',
    'run': '按审核规则执行结构化排版',
    'toc': '为单节正文创建原生目录',
    'paginate': '单节目录稿分节并设置正文页码',
    'flow': '图表同页、一级标题分节、独立页眉页脚',
    'chapter': '指定章节与层级的局部样式修改',
    'template-info': '读取模板样式与分节',
    'layout': '按明确节映射迁移页面设置',
    'compare': '生成 DOCX 差异报告',
    'repair-toc': '保守修复目录链接',
    'reviewed-template': '执行与原稿哈希绑定的四节模板审核清单',
    'crossrefs': '手打引用转换为REF，参考文献转换为原生列表',
}


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = argparse.ArgumentParser(
        description='PaperFit 本地论文格式工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='命令\n' + '\n'.join(f'  {name:20} {help}' for name, help in COMMANDS.items())
        + '\n\n使用 COMMAND --help 查看参数。修改命令要求新的输出路径。')
    parser.add_argument('--version', action='version', version='PaperFit ' +
                        (Path(__file__).resolve().parents[1]/'VERSION').read_text().strip())
    parser.add_argument('command', nargs='?', choices=COMMANDS)
    if not argv or argv[0].startswith('-'):
        parser.parse_args(argv)
        if not argv:
            parser.print_help()
        return
    if argv[0] not in COMMANDS:
        parser.error('未知命令：' + argv[0])
    try:
        if argv[0] == 'crossrefs':
            sub=argparse.ArgumentParser(prog='paperfit.py crossrefs',description=COMMANDS[argv[0]])
            sub.add_argument('source',type=Path);sub.add_argument('output',type=Path)
            args=sub.parse_args(argv[1:])
            from native_references import convert
            report=convert(args.source,args.output)
            print(report['status']+'  '+str(args.output))
        elif argv[0] in ('extract', 'audit', 'run'):
            from phase2 import main as dispatch
            dispatch(argv)
        elif argv[0] in ('chapter', 'template-info', 'layout', 'compare', 'repair-toc'):
            from phase3 import main as dispatch
            dispatch(argv)
        elif argv[0] == 'reviewed-template':
            sub = argparse.ArgumentParser(prog='paperfit.py reviewed-template', description=COMMANDS[argv[0]])
            sub.add_argument('source', type=Path)
            sub.add_argument('manifest', type=Path)
            sub.add_argument('output', type=Path)
            args = sub.parse_args(argv[1:])
            from reviewed_template import format_template
            report = format_template(args.source, json.loads(args.manifest.read_text(encoding='utf-8')), args.output)
            print(('needs_review' if report['pending'] else 'requires_layout_check') + '  ' + str(args.output))
        else:
            from paperfit import main as dispatch
            dispatch(argv)
    except (ValueError, KeyError, OSError) as error:
        parser.exit(2, f'PaperFit: {error}\n')


if __name__ == '__main__':
    main()
