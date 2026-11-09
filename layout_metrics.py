"""Read-only layout calculations from OOXML, independent of the HTML renderer.

Twips / 20 = points; EMUs / 12700 = points; CSS pixels = points * 96 / 72.
Line multipliers are in 240ths of a font's normal line height, not its em size.
"""
import io,re,zipfile
from decimal import Decimal
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def tag(name):return '{'+W+'}'+name
def attrs(node):return {k.rsplit('}',1)[-1]:v for k,v in node.attrib.items()} if node is not None else {}
def points(value):return float(Decimal(str(value))/20)
def on(node):return node is not None and node.get(tag('val'),'1') not in ('0','false','off')
def parse(data):return etree.fromstring(data,etree.XMLParser(resolve_entities=False,no_network=True))

def analyze_layout(raw):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:parts={n:z.read(n) for n in z.namelist() if re.fullmatch(r'word/(document|header\d*|footer\d*|footnotes|endnotes|styles|settings)\.xml',n)}
    styles=parse(parts['word/styles.xml']) if 'word/styles.xml' in parts else etree.Element('styles')
    by_id={s.get(tag('styleId')):s for s in styles.findall(tag('style'))}
    default=next((k for k,s in by_id.items() if s.get(tag('type'))=='paragraph' and s.get(tag('default'))=='1'),None)
    defaults=styles.find('./'+tag('docDefaults')+'/'+tag('pPrDefault')+'/'+tag('pPr'))
    def chain(style,seen=None):
        seen=set() if seen is None else seen
        if style not in by_id or style in seen:return []
        seen.add(style);s=by_id[style];base=s.find(tag('basedOn'))
        return chain(base.get(tag('val')) if base is not None else None,seen)+[s.find(tag('pPr'))]
    result={'sections':{},'paragraphs':{},'settings':{}}
    settings=parse(parts.get('word/settings.xml',b'<settings/>'))
    result['settings']={k:on(settings.find(tag(k))) for k in ['mirrorMargins','gutterAtTop']}
    feature=settings.find('./'+tag('compat')+'/'+tag('compatSetting')+'[@'+tag('name')+'="enableOpenTypeFeatures"]')
    result['settings']['openTypeFeatures']=on(feature)
    for name,data in parts.items():
        if not re.fullmatch(r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml',name):continue
        root=parse(data)
        for section in root.iter(tag('sectPr')):
            size=attrs(section.find(tag('pgSz')));margins=attrs(section.find(tag('pgMar')))
            result['sections']['s'+str(len(result['sections']))]={
                'specialHeaders':on(section.find(tag('titlePg'))) or any(e.get(tag('type')) in ('first','even') for e in section if e.tag in (tag('headerReference'),tag('footerReference'))),
                'widthPt':points(size['w']) if 'w' in size else None,
                'heightPt':points(size['h']) if 'h' in size else None,
                'marginsPt':{k:points(v) for k,v in margins.items() if k in ('top','bottom','left','right','gutter','header','footer')}}
        for p in root.iter(tag('p')):
            direct=p.find(tag('pPr'));style=direct.find(tag('pStyle')) if direct is not None else None
            style_id=style.get(tag('val')) if style is not None else default
            spacing={};indents={};contextual=None;flow={}
            for props in [defaults,*chain(style_id),direct]:
                if props is None:continue
                spacing.update(attrs(props.find(tag('spacing'))));indents.update(attrs(props.find(tag('ind'))))
                for key in ['keepNext','keepLines','widowControl','pageBreakBefore']:
                    node=props.find(tag(key))
                    if node is not None:flow[key]=on(node)
                context=props.find(tag('contextualSpacing'))
                if context is not None:contextual=on(context)
            item={'flow':flow,'style':style_id,'contextualSpacing':contextual,'indentsPt':{}}
            for key in ['before','after']:
                if key in spacing:item[key+'Pt']=points(spacing[key])
            if 'line' in spacing:
                rule=spacing.get('lineRule','auto');item['lineRule']=rule
                item['line']=float(Decimal(spacing['line'])/240) if rule=='auto' else points(spacing['line'])
            for key in ['left','right','start','end','firstLine','hanging']:
                if key in indents:item['indentsPt'][key]=points(indents[key])
            result['paragraphs']['p'+str(len(result['paragraphs']))]=item
    return result
