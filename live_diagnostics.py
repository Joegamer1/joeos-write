"""Private, text-free diagnostics for comparing ordinary and test launches."""
import hashlib,json,os,sys,time,winreg
from pathlib import Path
from PySide6.QtGui import QFont,QFontInfo,QFontDatabase

def write_diagnostics(window,browser):
    root=Path(sys.executable).parents[2] if getattr(sys,'frozen',False) else Path(__file__).parent
    folder=root/'.audit-private/launch-diagnostics';folder.mkdir(parents=True,exist_ok=True)
    registry=[]
    for label,hive in [('user',winreg.HKEY_CURRENT_USER),('machine',winreg.HKEY_LOCAL_MACHINE)]:
        try:
            with winreg.OpenKey(hive,r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts') as key:
                for i in range(winreg.QueryInfoKey(key)[1]):
                    name,value,_=winreg.EnumValue(key,i)
                    if 'aptos' in name.lower():registry.append({'scope':label,'name':name,'path':value})
        except OSError:pass
    fonts={name:{'listed':name in QFontDatabase.families(),'resolved':QFontInfo(QFont(name)).family()} for name in ['Aptos','Aptos Serif','Arial','Times New Roman','Noto Serif']}
    data={'time':time.time(),'pid':os.getpid(),'executable':sys.executable,'frozen':bool(getattr(sys,'frozen',False)),
          'local_app_data':os.environ.get('LOCALAPPDATA'),'source_path':window.path,
          'source_sha256':hashlib.sha256(window.model.original).hexdigest(),'document_fonts':sorted(window.model.fonts),
          'qt_fonts':fonts,'aptos_registry':registry,'browser':browser}
    target=folder/(str(os.getpid())+'.json');temporary=target.with_suffix('.tmp')
    temporary.write_text(json.dumps(data,indent=2),encoding='utf-8');os.replace(temporary,target)
