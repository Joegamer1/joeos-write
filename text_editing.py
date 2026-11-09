"""Range edits patch touched OOXML paragraphs, never rebuild the document."""
import re
from xml.sax.saxutils import escape
from preserved_docx import PreservedDocx, wn, scan, patch_bytes
from package_storage import rewrite_package


def empty_paragraphs(model):
    result=[];index=0
    for part,nodes in model.nodes.items():
        for node in nodes:
            if node.name!=wn('p'):continue
            if not node.locked and not model.readonly and all(n.name==wn('pPr') or (n.name==wn('r') and all(c.name==wn('rPr') for c in n.children)) for n in node.children):
                result.append({'id':'e'+str(index),'paragraph':'p'+str(index),'part':part,'start':node.start,'text':'','editable':True})
            index+=1
    return result


def utf16_offset(text, offset):
    raw = text.encode('utf-16-le')
    if not isinstance(offset, int) or offset < 0 or offset*2 > len(raw):
        raise ValueError('The cursor position is stale.')
    return len(raw[:offset*2].decode('utf-16-le'))


def paragraph(node):
    while node and node.name != wn('p'):
        node = node.parent
    if node is None:
        raise ValueError('This text is outside an editable paragraph.')
    return node


def endpoint(model, point):
    identity = point['id']
    if identity not in model.texts:
        raise ValueError('The cursor position is stale.')
    part, node = model.texts[identity]
    if node.locked or model.readonly:
        raise ValueError('This field or protected text cannot be edited.')
    offset = utf16_offset(node.text, point['offset'])
    return part, node, offset, paragraph(node)


def text_xml(tag, value):
    return ('<'+tag+' xml:space="preserve">'+escape(value)+'</'+tag+'>').encode()


def split_inner(data, node, offset, para):
    """Split at a text offset, closing/reopening original inline ancestors."""
    ancestors = []
    parent = node.parent
    while parent is not para:
        if parent is None or parent.name not in (wn('r'), wn('hyperlink'), wn('smartTag')):
            raise ValueError('This embedded structure is preserved but cannot be split.')
        ancestors.append(parent)
        parent = parent.parent
    left = data[para.head_end:node.start]+text_xml(node.tag, node.text[:offset])
    left += b''.join(data[n.end:n.outer_end] for n in ancestors)
    right = b''
    for n in reversed(ancestors):
        right += data[n.start:n.head_end]
        props = next((c for c in n.children if c.name == wn('rPr')), None)
        if props:
            right += data[props.start:props.outer_end]
    right += text_xml(node.tag, node.text[offset:])+data[node.outer_end:para.end]
    return left, right


def validate_span(current, a, b):
    pa, na, oa, para = endpoint(current, a)
    pb, nb, ob, endpara = endpoint(current, b)
    if pa != pb or para.parent is not endpara.parent:
        raise ValueError('Select text within one document story or table cell.')
    if (na.start, oa) > (nb.start, ob):
        raise ValueError('The selection order is invalid.')
    between = [n for n in para.parent.children if para.start <= n.start <= endpara.start]
    if any(n.name != wn('p') for n in between):
        raise ValueError('This selection crosses a table or embedded object.')
    protected = {wn(x) for x in ('drawing','pict','object','fldChar','sdt','ins','del','sectPr')}
    if any(na.start < n.start < nb.start and (n.locked or n.name in protected) for n in current.nodes[pa]):
        raise ValueError('This selection crosses a protected field, image, or section boundary.')
    return pa, na, oa, para, nb, ob, endpara


def replace_range(model, a, b, value, marks=None):
    if not isinstance(value, str) or len(value) > 1_000_000:
        raise ValueError('Text is too large.')
    if any((ord(c)<32 and c not in '\n\t\v\f') or 0xD800 <= ord(c) <= 0xDFFF or ord(c) in (0xFFFE,0xFFFF) for c in value):
        raise ValueError('Unsupported control character.')
    current = PreservedDocx(model.bytes())
    if a['id'].startswith('e') or b['id'].startswith('e'):
        if a!=b or a['offset']!=0:
            raise ValueError('This selection includes an empty structural paragraph. Place the cursor inside it to type.')
        empty=next((p for p in empty_paragraphs(current) if p['id']==a['id']),None)
        if empty is None:raise ValueError('The empty paragraph position is stale.')
        part=empty['part'];data=current.parts[part]
        node=next(n for n in current.nodes[part] if n.start==empty['start'])
        prefix=node.tag.rsplit(':',1)[0]+':' if ':' in node.tag else ''
        opening=data[node.start:node.head_end]
        if node.empty:opening=opening[:-2]+b'>'
        inner=b'' if node.empty else data[node.head_end:node.end]
        mark=opening+inner+('<'+prefix+'r>').encode()
        at=node.start+len(mark)
        replacement=mark+text_xml(prefix+'t','')+('</'+prefix+'r></'+node.tag+'>').encode()
        current=PreservedDocx(rewrite_package(current.original,{part:data[:node.start]+replacement+data[node.outer_end:]}))
        identity=next(k for k,(p,n) in current.texts.items() if p==part and n.start==at)
        return replace_range(current,{'id':identity,'offset':0},{'id':identity,'offset':0},value,marks)
    part, start, off, para, end, endoff, endpara = validate_span(current, a, b)
    data = current.parts[part]
    left, _ = split_inner(data, start, off, para)
    _, right = split_inner(data, end, endoff, endpara)
    props = next((n for n in para.children if n.name == wn('pPr')), None)
    pprops = data[props.start:props.outer_end] if props else b''
    if '\n' in value and props and any(n.name == wn('sectPr') for n in props.children):
        raise ValueError('Splitting a section-ending paragraph is not yet supported.')
    prefix = start.tag.rsplit(':',1)[0]+':' if ':' in start.tag else ''
    run = start.parent
    rprops = next((n for n in run.children if n.name == wn('rPr')), None)
    rprops = data[rprops.start:rprops.outer_end] if rprops else b''
    if marks:
        # Reuse the validated formatting implementation and original properties.
        for key, val in marks.items():
            current.set_format(a['id'], key, val)
        rprops = current._format_patch(part, run, marks)[2]
    chunks = value.split('\n')
    out = data[para.start:para.head_end]+left
    caret_at = None
    caret_offset = 0
    for i, chunk in enumerate(chunks):
        if i:
            out += ('</'+para.tag+'><'+para.tag+'>').encode()+pprops
        out += ('<'+prefix+'r>').encode()+rprops
        tokens = re.split('([\t\v\f])', chunk)
        for token in tokens:
            if token in ('\t','\v','\f'):
                out += (('<'+prefix+'tab/>') if token=='\t' else ('<'+prefix+'br'+(' '+prefix+'type="page"' if token=='\f' else '')+'/>')).encode()
            else:
                caret_at = para.start+len(out)
                caret_offset = len(token.encode('utf-16-le'))//2
                out += text_xml(start.tag, token)
        out += ('</'+prefix+'r>').encode()
    out += right+data[endpara.end:endpara.outer_end]
    changed = data[:para.start]+out+data[endpara.outer_end:]
    result = PreservedDocx(rewrite_package(current.original,{part:changed}))
    identity = next(k for k,(p,n) in result.texts.items() if p==part and n.start==caret_at)
    return result, {'id':identity,'offset':caret_offset}


def format_range(model, a, b, marks):
    current = PreservedDocx(model.bytes())
    part, start, off, para, end, endoff, endpara = validate_span(current,a,b)
    # Apply from the end to keep earlier text IDs stable when splitting runs.
    targets = [(key,n) for key,(p,n) in current.texts.items() if p==part and start.start <= n.start <= end.start]
    for key, node in reversed(targets):
        lo = off if node is start else 0
        hi = endoff if node is end else len(node.text)
        if lo==hi:
            continue
        current = PreservedDocx(current.bytes())
        p,n = current.texts[key]
        run=n.parent; data=current.parts[p]
        if run.name!=wn('r'):
            raise ValueError('Unsupported formatting selection.')
        for prop,val in marks.items(): current.set_format(key,prop,val)
        newprops=current._format_patch(p,run,marks)[2]
        oldprops=next((c for c in run.children if c.name==wn('rPr')),None)
        properties=data[oldprops.start:oldprops.outer_end] if oldprops else b''
        opening=data[run.start:run.head_end];closing=data[run.end:run.outer_end]
        left=data[run.head_end:n.start]+text_xml(n.tag,n.text[:lo])
        right=properties+text_xml(n.tag,n.text[hi:])+data[n.outer_end:run.end]
        replacement=opening+left+closing+opening+newprops+text_xml(n.tag,n.text[lo:hi])+closing+opening+right+closing
        # Only this run is replaced; other part bytes and package entries stay exact.
        current=PreservedDocx(rewrite_package(current.original,{p:data[:run.start]+replacement+data[run.outer_end:]}))
    return current


def remap_format_selection(before, after, a, b):
    """Formatting splits runs, so map selection by unchanged story text offsets."""
    before=PreservedDocx(before.bytes())
    def position(point):
        part,node=before.texts[point['id']];offset=point['offset']
        for p,n in before.texts.values():
            if p==part and n.start<node.start:offset+=len(n.text.encode('utf-16-le'))//2
        return part,offset
    def locate(point,bias):
        part,left=position(point);last=None
        for key,(p,n) in after.texts.items():
            if p!=part:continue
            length=len(n.text.encode('utf-16-le'))//2;last={'id':key,'offset':length}
            if left<length or (left==length and bias=='end' and length):return {'id':key,'offset':left}
            left-=length
        return last
    return {'start':locate(a,'start'),'end':locate(b,'end')}


def paragraph_format(model, a, b, command):
    """Patch only requested paragraph properties, preserving other XML bytes."""
    from layout_metrics import analyze_layout
    current = PreservedDocx(model.bytes())
    layout = analyze_layout(current.bytes())['paragraphs']
    paragraphs = [(part, n) for part, nodes in current.nodes.items() for n in nodes if n.name == wn('p')]
    def locate(point):
        if point['id'].startswith('e'):
            entry = next(x for x in empty_paragraphs(current) if x['id'] == point['id'])
            return next((part,n) for part,n in paragraphs if part == entry['part'] and n.start == entry['start'])
        part,node,offset,p = endpoint(current,point)
        return part,p
    part,first = locate(a); endpart,last = locate(b)
    if part != endpart or first.parent is not last.parent:
        raise ValueError('Select paragraphs within one document story or table cell.')
    patches=[]
    for index,(name,p) in enumerate(paragraphs):
        if name != part or not first.start <= p.start <= last.start:continue
        if p.locked or current.readonly:raise ValueError('This paragraph is protected.')
        # A selection ending at the next paragraph start does not format it.
        if p is last and p is not first and b['offset']==0:continue
        data=current.parts[part];prefix=p.tag.rsplit(':',1)[0]+':' if ':' in p.tag else ''
        props=next((n for n in p.children if n.name==wn('pPr')),None)
        values={}; tag=None
        if command in ('left','center','right','both'):tag='jc';values={'val':command}
        elif command in ('single','double','onehalf'):tag='spacing';values={'line':str({'single':240,'double':480,'onehalf':360}[command]),'lineRule':'auto'}
        elif command in ('indent','outdent','hang','unhang'):
            tag='ind';effective=layout.get('p'+str(index),{}).get('indentsPt',{})
            left=round(effective.get('start',effective.get('left',0))*20)
            if command in ('indent','outdent'):
                values={'left':str(max(0,left+(720 if command=='indent' else -720)))}
                if 'start' in effective:values['start']=values['left']
            else:
                old=round(effective.get('hanging',0)*20);new=max(0,old+(720 if command=='hang' else -720))
                values={'hanging':str(new),'firstLine':'0','left':str(max(0,left+new-old))}
                if 'start' in effective:values['start']=values['left']
        else:raise ValueError('Unknown paragraph command.')
        child=next((n for n in props.children if n.name==wn(tag)),None) if props else None
        attributes=dict(child.attrs) if child else {}
        attributes.update({wn(k):v for k,v in values.items()})
        # Existing OOXML attributes retained, including spacing before/after.
        attributes={k:v for k,v in attributes.items() if k.startswith(wn(''))}
        xml=('<'+prefix+tag+''.join(' '+prefix+k.split('|')[-1]+'="'+escape(v,{'"':'&quot;'})+'"' for k,v in attributes.items())+'/>').encode()
        if child:patches.append((child.start,child.outer_end,xml))
        elif props and not props.empty:
            order=['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange']
            at=next((n.start for n in props.children if n.name.split('|')[-1] in order and order.index(n.name.split('|')[-1])>order.index(tag)),props.end)
            patches.append((at,at,xml))
        elif props:patches.append((props.start,props.outer_end,('<'+prefix+'pPr>').encode()+xml+('</'+prefix+'pPr>').encode()))
        elif p.empty:patches.append((p.start,p.outer_end,data[p.start:p.head_end-2]+b'>'+('<'+prefix+'pPr>').encode()+xml+('</'+prefix+'pPr></'+p.tag+'>').encode()))
        else:patches.append((p.head_end,p.head_end,('<'+prefix+'pPr>').encode()+xml+('</'+prefix+'pPr>').encode()))
    return PreservedDocx(rewrite_package(current.original,{part:patch_bytes(current.parts[part],patches)}))
