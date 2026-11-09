"""Build and verify the portable archive without private documents or preferences."""
import json,os,shutil,zipfile
from pathlib import Path
root=Path('dist-keyboard/JoeOS Write')
for name in ['README.md','VERIFICATION.md']:shutil.copyfile(name,root/name)
shutil.copytree('THIRD-PARTY-LICENSES',root/'THIRD-PARTY-LICENSES',dirs_exist_ok=True)
temporary=Path('.tmp/joeos-portable.zip');archive=Path('JoeOS-Write-Windows.zip')
with zipfile.ZipFile(temporary,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for p in root.rglob('*'):
        if p.is_file():z.write(p,Path('JoeOS Write')/p.relative_to(root))
with zipfile.ZipFile(temporary) as z:
    assert z.testzip() is None
    assert [n for n in z.namelist() if n.lower().endswith('.docx')]==['JoeOS Write/_internal/docx/templates/default.docx']
    assert not any(n.lower().endswith(('.ttf','.otf','.ttc','.otc')) or 'preview-profiles.json' in n for n in z.namelist())
for p in Path('assets/editor').iterdir():
    if p.is_file():assert p.read_bytes()==(root/'_internal/assets/editor'/p.name).read_bytes()
os.replace(temporary,archive)
report={'zip_bytes':archive.stat().st_size,'zip_valid':True,'personal_documents_included':False,'microsoft_fonts_bundled':False,'preview_preferences_bundled':False,'editor_assets_match_source':True}
Path('audit-output/portable-verification.json').write_text(json.dumps(report,indent=2));print(json.dumps(report))
