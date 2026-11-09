"""Original-package DOCX editing. Never rebuild a document from its rendered HTML."""
from __future__ import annotations
import copy, io, math, os, re, tempfile, zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.parsers import expat
from xml.sax.saxutils import escape
from lxml import etree
from package_storage import PackageParts, rewrite_package
from media_preview import preview_media

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML='http://www.w3.org/XML/1998/namespace'
NS={'w':W}
def wn(s):return W+'|'+s

@dataclass
class Element:
    name:str
    tag:str
    start:int
    head_end:int
    attrs:dict
    parent:object=None
    end:int=0
    outer_end:int=0
    empty:bool=False
    text:str=''
    locked:bool=False
    children:list=field(default_factory=list)

def tag_end(data,start):
    quote=None
    for i in range(start,len(data)):
        b=data[i]
        if quote:
            if b==quote:quote=None
        elif b in (34,39):quote=b
        elif b==62:return i+1
    raise ValueError('Unterminated XML tag.')

def scan(data):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():raise ValueError('DTD/entity declarations are not supported.')
    if data.startswith((b'\xff\xfe',b'\xfe\xff')):raise ValueError('UTF-16 XML is not yet supported; original remains unchanged.')
    parser=expat.ParserCreate(namespace_separator='|');stack=[];all_nodes=[];field_depth=0
    def start(name,attrs):
        nonlocal field_depth
        pos=parser.CurrentByteIndex;end=tag_end(data,pos);head=data[pos:end]
        tag=re.match(rb'<([^\s/>]+)',head).group(1).decode('utf-8')
        if name==wn('fldChar') and attrs.get(wn('fldCharType'))=='begin':field_depth+=1
        parent=stack[-1] if stack else None
        locked=bool(field_depth or (parent and parent.locked) or name in {wn(x) for x in ['ins','del','moveFrom','moveTo','sdt','fldSimple','txbxContent']})
        n=Element(name,tag,pos,end,attrs,parent,empty=head.rstrip().endswith(b'/>'),locked=locked)
        if parent:parent.children.append(n)
        stack.append(n);all_nodes.append(n)
    def end(name):
        nonlocal field_depth
        n=stack.pop();n.end=n.head_end if n.empty else parser.CurrentByteIndex
        n.outer_end=n.head_end if n.empty else tag_end(data,n.end)
        if name==wn('fldChar') and n.attrs.get(wn('fldCharType'))=='end':field_depth=max(0,field_depth-1)
    def text(value):
        if stack:stack[-1].text+=value
    parser.StartElementHandler=start;parser.EndElementHandler=end;parser.CharacterDataHandler=text
    parser.Parse(data,True)
    return all_nodes

def patch_bytes(data,patches):
    patches=sorted(patches,key=lambda p:p[0],reverse=True);last=len(data)+1
    for start,end,replacement in patches:
        if end>last:raise ValueError('Overlapping document changes were rejected.')
        data=data[:start]+replacement+data[end:];last=start
    return data

def valid_text(value):
    if not isinstance(value,str) or len(value)>1_000_000:raise ValueError('Text is too large.')
    if any((ord(c)<32 or 0xD800<=ord(c)<=0xDFFF or ord(c) in (0xFFFE,0xFFFF)) for c in value):
        raise ValueError('Use Add paragraph for a new paragraph. Tabs, line breaks and control characters cannot be inserted into this text segment.')

class PreservedDocx:
    def __init__(self,raw:bytes):
        self.original=raw;self.text_edits={};self.format_edits={};self.parts={};self.nodes={};self.texts={};self.runs={};self.readonly=False
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            infos=z.infolist()
            if len(infos)>10000 or sum(i.file_size for i in infos)>800_000_000 or sum(i.file_size for i in infos if not (i.filename.startswith('word/media/') and i.filename.lower().endswith('.wmf')))>100_000_000 or any(i.file_size>512_000_000 for i in infos):raise ValueError('Document exceeds the safe processing limit.')
            if len({i.filename for i in infos})!=len(infos):raise ValueError('Duplicate package entries are not supported.')
            if any(i.flag_bits&1 for i in infos):raise ValueError('Encrypted documents are not supported.')
            if 'word/document.xml' not in z.namelist():raise ValueError('Not a supported Word DOCX document.')
            self.parts=PackageParts(raw)
        self._preview_media=None;self.preview_notices=[]
        self.readonly=any(n.startswith('_xmlsignatures/') for n in self.parts)
        settings=self.parts.get('word/settings.xml',b'')
        if settings:
            for node in scan(settings):
                if node.name==wn('documentProtection') and node.attrs.get(wn('enforcement'),'0') not in ('0','false','off'):self.readonly=True
        self.fonts=set();self.unsupported=set()
        for name in self.parts:
            if name.endswith('.xml') and name.startswith('word/'):
                data=self.parts[name] if re.fullmatch(r'word/(document|header\d*|footer\d*|footnotes|endnotes|styles|fontTable)\.xml',name) else b''
                # Parse only document stories and font/style metadata; other parts remain opaque.
                if re.fullmatch(r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml',name):
                    nodes=scan(data);self.nodes[name]=nodes
                    for n in nodes:
                        if n.name==wn('r'):
                            self.runs[(name,n.start)]=n
                        if n.name==wn('t') and n.parent and n.parent.name==wn('r') and not n.children:
                            identity='t'+str(len(self.texts));self.texts[identity]=(name,n)
                        if n.name in {wn(x) for x in ['sdt','ins','del','fldChar','fldSimple','txbxContent','altChunk']}:
                            self.unsupported.add(n.name.split('|')[-1])
                if name in ('word/styles.xml','word/document.xml','word/fontTable.xml'):
                    for n in scan(data):
                        if n.name==wn('rFonts'):
                            self.fonts.update(v for k,v in n.attrs.items() if k in {wn('ascii'),wn('hAnsi'),wn('eastAsia'),wn('cs')} and v)

    @classmethod
    def open(cls,path):return cls(Path(path).read_bytes())

    @classmethod
    def blank(cls):
        from docx import Document
        from docx.shared import Inches
        d=Document();s=d.sections[0];s.top_margin=s.bottom_margin=s.left_margin=s.right_margin=Inches(1)
        p=d.add_paragraph();r=p.add_run();t=etree.SubElement(r._r,'{'+W+'}t');t.text=''
        out=io.BytesIO();d.save(out);return cls(out.getvalue())

    def snapshot(self):return self.original,dict(self.text_edits),copy.deepcopy(self.format_edits)

    @classmethod
    def from_snapshot(cls,state):
        d=cls(state[0]);d.text_edits=dict(state[1]);d.format_edits=copy.deepcopy(state[2]);return d

    def metadata(self):
        return [{'id':key,'text':self.text_edits.get(key,n.text),'part':part,'editable':not(n.locked or self.readonly),
                 'run':part+':'+str(n.parent.start)} for key,(part,n) in self.texts.items()]

    def set_text(self,identity,value):
        valid_text(value)
        if identity not in self.texts:raise ValueError('Text selection is stale. Reopen the document.')
        part,n=self.texts[identity]
        if self.readonly or n.locked:raise ValueError('This field, revision, content control or signed/protected document is preserved but not editable here.')
        if value==n.text:self.text_edits.pop(identity,None)
        else:self.text_edits[identity]=value

    def set_format(self,identity,key,value):
        if identity not in self.texts:raise ValueError('No editable text selected.')
        part,n=self.texts[identity]
        if self.readonly or n.locked:raise ValueError('This text is protected.')
        if key not in ('bold','italic','underline','font','size','color'):raise ValueError('Unsupported formatting operation.')
        if key in ('bold','italic','underline') and not isinstance(value,bool):raise ValueError('Invalid toggle.')
        if key=='size':
            value=float(value)
            if not math.isfinite(value) or not 1<=value<=400 or value*2!=int(value*2):raise ValueError('Use a size from 1 to 400 in half-point steps.')
        if key=='font' and (not isinstance(value,str) or not value.strip() or len(value)>128):raise ValueError('Invalid font name.')
        if key=='color' and not re.fullmatch(r'#[0-9a-fA-F]{6}',str(value)):raise ValueError('Invalid color.')
        self.format_edits.setdefault((part,n.parent.start),{})[key]=value

    def _format_patch(self,part,run,changes):
        old=next((c for c in run.children if c.name==wn('rPr')),None)
        prefix=run.tag.rsplit(':',1)[0] if ':' in run.tag else None
        # Parse the run property subtree in the original namespace context.
        root=etree.fromstring(self.parts[part],etree.XMLParser(resolve_entities=False,no_network=True))
        parsed_runs=list(root.iter('{'+W+'}r'));run_index=next(i for i,n in enumerate(n for n in self.nodes[part] if n.name==wn('r')) if n is run)
        source=parsed_runs[run_index].find('{'+W+'}rPr')
        props=copy.deepcopy(source) if source is not None else etree.Element('{'+W+'}rPr',nsmap={prefix:W})
        def property(name,attrs,remove=()):
            tag='{'+W+'}'+name;el=props.find(tag)
            if el is None:el=etree.SubElement(props,tag)
            for attr in remove:el.attrib.pop('{'+W+'}'+attr,None)
            for attr,value in attrs.items():el.set('{'+W+'}'+attr,str(value))
        for key,value in changes.items():
            if key=='bold':property('b',{'val':'1' if value else '0'})
            elif key=='italic':property('i',{'val':'1' if value else '0'})
            elif key=='underline':property('u',{'val':'single' if value else 'none'})
            elif key=='font':property('rFonts',{'ascii':value,'hAnsi':value},('asciiTheme','hAnsiTheme'))
            elif key=='size':
                property('sz',{'val':int(value*2)});property('szCs',{'val':int(value*2)})
            elif key=='color':property('color',{'val':value[1:].upper()},('themeColor','themeTint','themeShade'))
        replacement=etree.tostring(props,encoding='utf-8',with_tail=False)
        return (old.start,old.outer_end,replacement) if old else (run.head_end,run.head_end,replacement)

    def _changed_parts(self,preview=False):
        patches={}
        if preview:
            counts={'p':0,'sectPr':0}
            for part,nodes in self.nodes.items():
                for n in nodes:
                    kind=n.name.split('|')[-1]
                    if n.name not in (wn('p'),wn('sectPr')):continue
                    identity=('p' if kind=='p' else 's')+str(counts[kind]);counts[kind]+=1
                    attribute='data-joeos-paragraph' if kind=='p' else 'data-joeos-section'
                    at=n.head_end-(2 if n.empty else 1)
                    patches.setdefault(part,[]).append((at,at,(' '+attribute+'="'+identity+'"').encode()))
        for identity,(part,n) in self.texts.items():
            if identity not in self.text_edits and not preview:continue
            head=self.parts[part][n.start:n.head_end]
            value=self.text_edits.get(identity,n.text)
            if identity in self.text_edits:
                if re.search(rb'\bxml:space\s*=',head):head=re.sub(rb'xml:space\s*=\s*([\'"]).*?\1',b'xml:space="preserve"',head)
                else:head=head[:-2]+b' xml:space="preserve"/>' if n.empty else head[:-1]+b' xml:space="preserve">'
            if preview:
                attr=(' data-joeos-id="'+identity+'"').encode()
                head=head[:-2]+attr+b'/>' if n.empty else head[:-1]+attr+b'>'
            if n.empty:
                replacement=head[:-2]+b'>'+escape(value).encode()+('</'+n.tag+'>').encode()
            elif identity in self.text_edits:
                replacement=head+escape(value).encode()+self.parts[part][n.end:n.outer_end]
            else:replacement=head+self.parts[part][n.head_end:n.outer_end]
            patches.setdefault(part,[]).append((n.start,n.outer_end,replacement))
        for (part,pos),changes in self.format_edits.items():patches.setdefault(part,[]).append(self._format_patch(part,self.runs[(part,pos)],changes))
        return {part:patch_bytes(self.parts[part],p) for part,p in patches.items()}

    def bytes(self,preview=False):
        if not preview and not self.text_edits and not self.format_edits:return self.original
        changed=self._changed_parts(preview)
        for data in changed.values():scan(data)
        if preview:
            if self._preview_media is None:self._preview_media,self.preview_notices=preview_media(self.original)
            changed.update(self._preview_media)
        return rewrite_package(self.original,changed)

    def append_paragraph(self,identity=None):
        if self.readonly:raise ValueError('This document is protected.')
        # Commit pending edits into a new in-memory base, never into the source file.
        current=PreservedDocx(self.bytes());part='word/document.xml';nodes=current.nodes[part]
        body=next(n for n in nodes if n.name==wn('body'))
        target=None
        if identity and identity in current.texts:
            story,n=current.texts[identity]
            if story==part and not n.locked:
                target=n
                while target and target.name!=wn('p'):target=target.parent
                if target and target.parent is not body:target=None
        if target is None:
            candidates=[n for n in body.children if n.name==wn('p')];target=candidates[-1] if candidates else None
        prefix=body.tag.split(':')[0]+':' if ':' in body.tag else ''
        props=b''
        if target:
            prop=next((n for n in target.children if n.name==wn('pPr')),None)
            if prop and not any(c.name==wn('sectPr') for c in prop.children):props=current.parts[part][prop.start:prop.outer_end]
        markup=f'<{prefix}p>'.encode()+props+f'<{prefix}r><{prefix}t></{prefix}t></{prefix}r></{prefix}p>'.encode()
        at=target.outer_end if target else next((n.start for n in body.children if n.name==wn('sectPr')),body.end)
        data=current.parts[part]
        return PreservedDocx(rewrite_package(current.original,{part:data[:at]+markup+data[at:]}))

    def save(self,path):
        raw=self.bytes();PreservedDocx(raw)
        fd,tmp=tempfile.mkstemp(prefix='.joeos-',suffix='.docx',dir=Path(path).parent)
        try:
            with os.fdopen(fd,'wb') as f:f.write(raw);f.flush();os.fsync(f.fileno())
            os.replace(tmp,path)
        finally:
            if os.path.exists(tmp):os.unlink(tmp)
