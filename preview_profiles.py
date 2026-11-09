"""Local display preferences, kept outside document packages."""
import json,os,tempfile
from pathlib import Path

class PreviewProfiles:
    def __init__(self,path=None):
        self.path=Path(path) if path else Path(os.environ.get('LOCALAPPDATA',Path.home()))/'JoeOS'/'Write'/'preview-profiles.json'
    def _read(self):
        if not self.path.exists():return {}
        data=json.loads(self.path.read_text(encoding='utf-8'))
        if not isinstance(data,dict):raise ValueError('Invalid local preview preferences.')
        return data
    def get(self,key):return self._read().get(key,{})
    def put(self,key,profile):
        data=self._read()
        if profile:data[key]=profile
        else:data.pop(key,None)
        self.path.parent.mkdir(parents=True,exist_ok=True)
        fd,tmp=tempfile.mkstemp(prefix='.preview-',suffix='.json',dir=self.path.parent)
        try:
            with os.fdopen(fd,'w',encoding='utf-8') as f:
                json.dump(data,f,indent=2);f.flush();os.fsync(f.fileno())
            os.replace(tmp,self.path)
        finally:
            if os.path.exists(tmp):os.unlink(tmp)
