"""Build a local, allowlisted Skill source bundle without user documents."""
import argparse
from hashlib import sha256
import json
from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED


def build(destination):
    root=Path(__file__).resolve().parents[1]
    version=(root/'VERSION').read_text().strip()
    destination=Path(destination);destination.mkdir(parents=True,exist_ok=True)
    archive=destination/f'paperfit-v{version}.zip'
    checksum=archive.with_suffix('.zip.sha256')
    if archive.exists() or checksum.exists():raise ValueError('交付包或校验文件已存在，请使用新目录')
    files=[root/n for n in ('VERSION','README.md','SKILL.md','requirements.txt','requirements-dev.txt','requirements-tested.txt')]
    files+=sorted((root/'scripts').glob('*.py'))+sorted((root/'tests').glob('*.py'))+sorted((root/'references').glob('*.md'))
    manifest={'version':version,'files':{str(p.relative_to(root)):sha256(p.read_bytes()).hexdigest() for p in files}}
    with ZipFile(archive,'x',compression=ZIP_DEFLATED) as z:
        for path in files:
            info=ZipInfo('paperfit/'+path.relative_to(root).as_posix(),date_time=(2026,11,7,0,0,0))
            info.compress_type=ZIP_DEFLATED;info.external_attr=0o644 << 16
            z.writestr(info,path.read_bytes())
        info=ZipInfo('paperfit/MANIFEST.json',date_time=(2026,11,7,0,0,0));info.compress_type=ZIP_DEFLATED
        z.writestr(info,json.dumps(manifest,ensure_ascii=False,indent=2).encode('utf-8'))
    digest=sha256(archive.read_bytes()).hexdigest()
    with checksum.open('x',encoding='utf-8') as f:f.write(f'{digest}  {archive.name}\n')
    return archive


if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('destination',type=Path);args=p.parse_args()
    try:print(build(args.destination))
    except (ValueError,OSError) as error:p.exit(2,f'PaperFit: {error}\n')
