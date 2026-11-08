"""Create a synthetic fixture and exercise the public CLI, entirely locally."""
import argparse
import base64
from hashlib import sha256
import json
from pathlib import Path
import subprocess
import sys


def smoke(folder):
    from docx import Document
    from docx.shared import Inches
    from phase2 import read_package, xml, DOC
    from paperfit import q
    from structured import visible
    from word_fields import scan
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=False)
    source = folder/'input.docx'
    image = folder/'fixture.png'
    image.write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jh1sAAAAASUVORK5CYII='))
    doc = Document()
    doc.add_heading('第一章 合成样例', 1)
    doc.add_paragraph('这是仅用于运行验证的合成正文，数值为 123。')
    doc.add_heading('1.1 测试小节', 2)
    doc.add_picture(str(image), width=Inches(0.5))
    doc.add_paragraph('图7 合成测试图')
    doc.add_paragraph('表8 合成测试表')
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = '项目'
    table.cell(0, 1).text = '数值'
    table.cell(1, 0).text = '合成数据'
    table.cell(1, 1).text = '123'
    doc.add_heading('第二章 检查', 1)
    doc.add_paragraph('此样例不包含用户论文或真实研究数据。')
    doc.save(source)
    digest = sha256(source.read_bytes()).hexdigest()
    plan = {'source_sha256':digest,'pending':[],
            'styles':{'body':{'font':'宋体','size':12,'line':1,'indent_chars':2},
                      'heading1':{'font':'宋体','size':15,'bold':True,'line':1},
                      'heading2':{'font':'宋体','size':14,'bold':True,'line':1}},
            'paragraph_roles':{'0':'heading1','1':'body','2':'heading2','6':'heading1','7':'body'},
            'structure':{'captions':'continuous'}}
    rules = folder/'rules.json'
    rules.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding='utf-8')
    cli = Path(__file__).with_name('paperfit.py').resolve()
    calls = [('audit',source,folder/'audit.json'),
             ('run',source,rules,folder/'formatted.docx'),
             ('toc',folder/'formatted.docx',folder/'toc.docx'),
             ('paginate',folder/'toc.docx',folder/'pages.docx'),
             ('flow',folder/'pages.docx',folder/'final.docx')]
    for args in calls:
        subprocess.run([sys.executable,str(cli),*map(str,args)],check=True,capture_output=True,text=True)
    final = folder/'final.docx'
    a,b = read_package(source),read_package(final)
    root = xml(b[DOC])
    assert sha256(source.read_bytes()).hexdigest() == digest
    assert all(b[k] == a[k] for k in a if k.startswith('word/media/'))
    assert [[visible(c) for c in t.iter(q('tc'))] for t in xml(a[DOC]).iter(q('tbl'))] == [[visible(c) for c in t.iter(q('tc'))] for t in root.iter(q('tbl'))]
    fields, errors = scan(root)
    assert not errors
    assert [f.value for f in fields if f.code.strip().startswith('SEQ ')] == ['1','1']
    assert len(list(root.iter(q('sectPr')))) == 3
    flow = json.loads(Path(str(final)+'.report.json').read_text())
    assert not flow['pending']
    # Repeated flow must not add further sections or mutate package parts.
    again = folder/'repeat.docx'
    subprocess.run([sys.executable,str(cli),'flow',str(final),str(again)],check=True,capture_output=True,text=True)
    assert read_package(again) == b
    refused = subprocess.run([sys.executable,str(cli),'flow',str(final),str(final)],capture_output=True,text=True)
    assert refused.returncode == 2 and read_package(final) == b
    report = {'status':'passed','source_unchanged':True,'media_unchanged':True,
              'table_contents_unchanged':True,'sections':3,'repeat_stable':True,
              'overwrite_rejected':True,'visual_acceptance':'not_performed_synthetic_smoke_only'}
    (folder/'smoke-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return report


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('new_output_directory',type=Path)
    args=parser.parse_args()
    try:
        print(json.dumps(smoke(args.new_output_directory),ensure_ascii=False))
    except subprocess.CalledProcessError as error:
        parser.exit(2,error.stderr or error.stdout or str(error))
    except (ValueError,OSError) as error:
        parser.exit(2,f'PaperFit: {error}\n')
