"""Bounded preview of the bitmap-only WMF dialect used by legacy Word imports.

Never executes metafile commands or OLE. The original member remains untouched.
"""
import io,struct,zlib,zipfile
from lxml import etree

def exact(stream,n):
    data=stream.read(n)
    if len(data)!=n:raise ValueError('Truncated image')
    return data

def skip(stream,n):
    while n:
        block=exact(stream,min(n,65536));n-=len(block)

def png_chunk(kind,data):
    return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data)&0xffffffff)

def wmf_bitmap_preview(stream,limit=1800):
    header=exact(stream,18)
    if struct.unpack('<HH',header[:4])!=(1,9):raise ValueError('Unsupported metafile header')
    rows=None;ow=oh=0
    for _ in range(10000):
        size,command=struct.unpack('<IH',exact(stream,6));remaining=size*2-6
        if remaining<0 or remaining>512_000_000:raise ValueError('Invalid metafile record')
        if command==0:break
        if command!=0x0b41:
            if command not in (0x103,0x20b,0x20c,0x1e,0x107,0x127):raise ValueError('Unsupported metafile drawing')
            skip(stream,remaining);continue
        if rows is not None or remaining<60:raise ValueError('Unsupported composite metafile')
        operation=exact(stream,20);dib=exact(stream,40);remaining-=60
        length,w,h,planes,bits,compression=struct.unpack('<IiiHHI',dib[:20])
        if length!=40 or planes!=1 or bits!=24 or compression!=0 or w<=0 or not h or max(w,abs(h))>50000:
            raise ValueError('Unsupported bitmap encoding')
        # Only complete, unscaled-source SRCCOPY bitmaps; reject crops/ROP effects.
        rop,sh,sw,sy,sx,dh,dw,dy,dx=struct.unpack('<I8h',operation)
        if rop!=0xcc0020 or (sw,sh,sx,sy)!=(w,abs(h),0,0):raise ValueError('Unsupported bitmap crop')
        stride=(w*3+3)&~3
        if stride*abs(h)>remaining:raise ValueError('Truncated bitmap')
        factor=max(1,w/limit,abs(h)/limit);ow=max(1,round(w/factor));oh=max(1,round(abs(h)/factor))
        wanted={min(abs(h)-1,int(y*abs(h)/oh)):y for y in range(oh)}
        rows=[None]*oh;xs=[min(w-1,int(x*w/ow))*3 for x in range(ow)]
        for y in range(abs(h)):
            row=exact(stream,stride);logical=abs(h)-1-y if h>0 else y
            if logical in wanted:
                rgb=bytearray(ow*3)
                for x,at in enumerate(xs):rgb[x*3:x*3+3]=row[at:at+3][::-1]
                rows[wanted[logical]]=b'\0'+rgb
        skip(stream,remaining-stride*abs(h))
    if rows is None:raise ValueError('Metafile has no supported bitmap')
    return b'\x89PNG\r\n\x1a\n'+png_chunk(b'IHDR',struct.pack('>IIBBBBB',ow,oh,8,2,0,0,0))+png_chunk(b'IDAT',zlib.compress(b''.join(rows),6))+png_chunk(b'IEND',b'')

def preview_media(raw):
    changed={};notices=[]
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        for info in z.infolist():
            if not info.filename.startswith('word/media/') or not info.filename.lower().endswith('.wmf'):continue
            try:
                with z.open(info) as stream:changed[info.filename]=wmf_bitmap_preview(stream)
                notices.append('Legacy image shown as a reduced-resolution preview; original image retained on save.')
            except ValueError:
                changed[info.filename]=b'<svg xmlns="http://www.w3.org/2000/svg" width="576" height="432"><rect width="100%" height="100%" fill="#eee6d8"/><text x="20" y="50">Legacy image preserved; preview unavailable</text></svg>'
                notices.append('A legacy image is preserved but its drawing format cannot be previewed.')
        if changed:
            ns='http://schemas.openxmlformats.org/package/2006/content-types'
            root=etree.fromstring(z.read('[Content_Types].xml'),etree.XMLParser(resolve_entities=False,no_network=True))
            for name,data in changed.items():
                for old in list(root):
                    if old.get('PartName')=='/'+name:root.remove(old)
                etree.SubElement(root,'{'+ns+'}Override',PartName='/'+name,ContentType='image/png' if data.startswith(b'\x89PNG') else 'image/svg+xml')
            changed['[Content_Types].xml']=etree.tostring(root)
    return changed,list(dict.fromkeys(notices))
