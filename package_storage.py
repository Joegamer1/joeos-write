"""Lazy package reads and compressed-byte preservation of untouched ZIP members."""
import copy, io, zipfile
from collections.abc import Mapping

class PackageParts(Mapping):
    def __init__(self, raw):
        self.raw=raw
        with zipfile.ZipFile(io.BytesIO(raw)) as z:self.names=tuple(z.namelist())
        self.cache={}
    def __iter__(self):return iter(self.names)
    def __len__(self):return len(self.names)
    def __getitem__(self, name):
        if name not in self.cache:
            with zipfile.ZipFile(io.BytesIO(self.raw)) as z:self.cache[name]=z.read(name)
        return self.cache[name]

def rewrite_package(raw, changed):
    """Copy local records verbatim, including descriptors; regenerate the directory.

    Python's zipfile writer owns the central directory/ZIP64 handling. Keeping this
    adapter isolated avoids inflating opaque media or changing its compression.
    """
    out=io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw)) as src,zipfile.ZipFile(out,'w') as dst:
        ordered=sorted(src.infolist(),key=lambda i:i.header_offset)
        ends={i.filename:(ordered[n+1].header_offset if n+1<len(ordered) else src.start_dir) for n,i in enumerate(ordered)}
        dst.comment=src.comment
        for original in src.infolist():
            info=copy.copy(original)
            if info.filename in changed:
                dst.writestr(info,changed[info.filename]);continue
            dst.fp.seek(dst.start_dir)
            info.header_offset=dst.fp.tell()
            dst.fp.write(memoryview(raw)[original.header_offset:ends[original.filename]])
            dst.start_dir=dst.fp.tell();dst.filelist.append(info);dst.NameToInfo[info.filename]=info;dst._didModify=True
    return out.getvalue()
