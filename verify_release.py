"""Verify packaged application code and assets against the current source."""
import hashlib,json,marshal,sys,types
from pathlib import Path
from PyInstaller.archive.readers import CArchiveReader

root=Path(sys.argv[1] if len(sys.argv)>1 else 'dist-keyboard/JoeOS Write')
exe=root/'JoeOS Write.exe';archive=CArchiveReader(str(exe));pyz=archive.open_embedded_archive('PYZ.pyz')
def normalized(code):
    return code.replace(co_filename='',co_consts=tuple(normalized(c) if isinstance(c,types.CodeType) else c for c in code.co_consts))
modules=['joeos_app','preserved_docx','layout_metrics','media_preview','package_storage','live_diagnostics','local_fonts','text_editing']
code_checks={}
for name in modules:
    packaged=marshal.loads(archive.extract(name)) if name=='joeos_app' else pyz.extract(name)
    source=compile(Path(name+'.py').read_bytes(),name+'.py','exec',optimize=0)
    code_checks[name]=normalized(packaged)==normalized(source)
assets={str(p):p.read_bytes()==(root/'_internal'/p).read_bytes() for p in Path('assets').rglob('*') if p.is_file()}
report={'root':str(root.resolve()),'exe_sha256':hashlib.sha256(exe.read_bytes()).hexdigest(),'code_matches':code_checks,'assets_checked':len(assets),'asset_mismatches':[p for p,ok in assets.items() if not ok]}
report['verified']=all(code_checks.values()) and all(assets.values())
Path('audit-output/release-sync.json').write_text(json.dumps(report,indent=2))
print(json.dumps(report,indent=2));sys.exit(0 if report['verified'] else 1)
