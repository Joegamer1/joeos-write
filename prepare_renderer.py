"""Pin and instrument the renderer; its output is never used as saved DOCX."""
from pathlib import Path
import shutil
out=Path('assets/editor');out.mkdir(parents=True,exist_ok=True)
source=Path('node_modules/docx-preview/dist/docx-preview.js').read_text(encoding='utf-8')
old='type: DomType.Text,\n                            text: c.textContent'
assert source.count(old)==1
source=source.replace(old,'type: DomType.Text,\n                            joeosId: c.getAttribute("data-joeos-id"),\n                            joeosPreserveSpace: (() => { for(let e=c;e;e=e.parentElement) { const v=e.getAttributeNS("http://www.w3.org/XML/1998/namespace","space"); if(v!==null)return v==="preserve"; } return false; })(),\n                            text: c.textContent')
old='renderText(elem) {\n            return this.h(elem.text);\n        }'
assert source.count(old)==1
source=source.replace(old,'''renderText(elem) {
            const text = elem.joeosPreserveSpace ? elem.text : elem.text.replace(/^[ \\t\\r\\n]+|[ \\t\\r\\n]+$/g, "");
            if (elem.joeosId) {
                const node = document.createElement("span");
                node.dataset.joeosId = elem.joeosId;
                node.dataset.joeosSpace = elem.joeosPreserveSpace ? "preserve" : "default";
                node.textContent = text;
                return node;
            }
            return this.h(text);
        }''')
def replace_once(old,new):
    global source
    assert source.count(old)==1, old
    source=source.replace(old,new)

# Word uses 240ths of a line when lineRule is omitted, not twentieths of a point.
replace_once('var result = { type: DomType.Paragraph, children: [] };',
             'var result = { type: DomType.Paragraph, children: [], joeosParagraphId: node.getAttribute("data-joeos-paragraph") };')
replace_once('var result = this.toHTML(elem, ns.html, "p");',
             'var result = this.toHTML(elem, ns.html, "p"); if(elem.joeosParagraphId) result.dataset.joeosParagraph = elem.joeosParagraphId;')
replace_once('var section = {};', 'var section = {joeosSectionId:elem.getAttribute("data-joeos-section")};')
replace_once('return this.h({ tagName: "section", className, style });',
             'const page = this.h({ tagName: "section", className, style }); if(props?.joeosSectionId) page.dataset.joeosSection = props.joeosSectionId; return page;')
replace_once('var num = parseInt(val) * usage.mul;', 'var num = parseFloat(val) * usage.mul;')
replace_once('Object.assign(result.style, style);', '''Object.assign(result.style, style);
                for (const [key,value] of Object.entries(style))
                    if (key.startsWith("--")) result.style.setProperty(key,value);''')
replace_once('var lineRule = globalXmlParser.attr(node, "lineRule");',
             'var lineRule = globalXmlParser.attr(node, "lineRule") || "auto";')
replace_once('style["line-height"] = `${(line / 240).toFixed(2)}`;',
             'style["--joeos-line-auto"] = `${line / 240}`; style["--joeos-line-min"] = "0pt"; style["line-height"] = "normal";')
replace_once('style["line-height"] = `calc(100% + ${line / 20}pt)`;',
             'style["--joeos-line-auto"] = "0"; style["--joeos-line-min"] = `${line / 20}pt`; style["line-height"] = "normal";')
replace_once('style["line-height"] = style["min-height"] = `${line / 20}pt`;',
             'style["--joeos-line-auto"] = "0"; style["--joeos-line-min"] = "0pt"; style["line-height"] = style["min-height"] = `${line / 20}pt`;')
# A gutter is additional binding space, separate from the ordinary page margins.
replace_once('case "evenAndOddHeaders":', '''case "mirrorMargins":
                case "gutterAtTop":
                    result[el.localName] = xml.boolAttr(el, "val", true);
                    break;
                case "evenAndOddHeaders":''')
replace_once('createPageElement(className, props, docStyle) {','createPageElement(className, props, docStyle, pageIndex = 0) {')
replace_once('this.createPageElement(this.className, props, document.cssStyle)',
             'this.createPageElement(this.className, props, document.cssStyle, i)')
replace_once('style.paddingBottom = props.pageMargins.bottom;', '''style.paddingBottom = props.pageMargins.bottom;
                    const settings = this.document.settingsPart?.settings || {};
                    if (settings.mirrorMargins && pageIndex % 2 === 1) {
                        [style.paddingLeft, style.paddingRight] = [style.paddingRight, style.paddingLeft];
                    }
                    const edge = settings.gutterAtTop ? "paddingTop" : settings.mirrorMargins && pageIndex % 2 === 1 ? "paddingRight" : "paddingLeft";
                    if (props.pageMargins.gutter) style[edge] = `calc(${style[edge] || '0pt'} + ${props.pageMargins.gutter})`;''')
# Paragraph border offsets are in points, not twips. Side padding/margin
# compensation is applied after render, preserving the original text indent.
replace_once('case "pBdr":\n                        this.parseBorderProperties(c, style);', '''case "pBdr":
                        this.parseBorderProperties(c, style);
                        for (const edge of globalXmlParser.elements(c)) {
                            if (["top", "bottom", "left", "right"].includes(edge.localName))
                                style["--joeos-border-space-" + edge.localName] = (globalXmlParser.intAttr(edge,"space",0)) + "pt";
                        }''')
replace_once('case "contextualSpacing":\n                    case "tblStyleColBandSize":',
             'case "tblStyleColBandSize":')
replace_once('case "kern":\n                        break;', '''case "contextualSpacing":
                        style["--joeos-context-spacing"] = globalXmlParser.boolAttr(c,"val",true) ? "1" : "0";
                        break;
                    case "kern":
                        style["--joeos-kern-pt"] = String(parseFloat(globalXmlParser.attr(c,"val")) / 2);
                        break;''')
replace_once('hyphens: auto; text-underline-position:', 'hyphens: manual; text-underline-position:')
replace_once('case "kern":\n                        style["--joeos-kern-pt"]', '''case "ligatures":
                        style["--joeos-ligatures"] = globalXmlParser.attr(c,"val");
                        break;
                    case "cntxtAlts":
                        style["--joeos-context-alternates"] = globalXmlParser.boolAttr(c,"val",true) ? "1" : "0";
                        break;
                    case "kern":
                        style["--joeos-kern-pt"]''')
replace_once('var isAnchor = node.localName == "anchor";', 'var isAnchor = node.localName == "anchor"; result.joeosInline = !isAnchor;')
replace_once('result.style.textIndent = "0px";', 'result.style.textIndent = "0px"; if(elem.joeosInline) result.dataset.joeosInlineImage = "true";')
# Match OOXML precedence: document defaults < paragraph style < character style
# < direct run formatting. The upstream default paragraph selector targets p
# rather than its runs, letting the default run font overwrite Normal's font.
replace_once('styleText += this.styleToString(selector, subStyle.values);', '''if (subStyle.target === "span") {
                        if (style.id == null) selector = `:where(.${this.className} span)`;
                        else if (style.target === "p") {
                            selector = `:where(p.${style.cssName}) span`;
                            if (defautStyles[style.target] === style)
                                selector = `:where(.${this.className} p), :where(.${this.className} p span), ` + selector;
                        } else if (style.target === "span" && defautStyles[style.target] === style)
                            selector = `:where(.${this.className} span), span.${style.cssName}`;
                    }
                    styleText += this.styleToString(selector, subStyle.values);''')
# Run character spacing is signed twips, distinct from paragraph spacing.
replace_once('this.parseSpacing(c, style);', 'this.parseSpacing(c, style);\n                        else if (elem.localName === "rPr")\n                            style["letter-spacing"] = globalXmlParser.lengthAttr(c, "val");')
# Embedded legacy objects expose only their static VML preview, never OLE.
replace_once('case "pict":\n                        result.children.push(this.parseVmlPicture(c));', 'case "object":\n                    case "pict":\n                        result.children.push(this.parseVmlPicture(c));')
replace_once('row.isHeader = globalXmlParser.boolAttr(c, "val");','row.isHeader = globalXmlParser.boolAttr(c, "val", true);')
replace_once('return this.toHTML(elem, ns.html, "tr", children);','const row = this.toHTML(elem, ns.html, "tr", children); if(elem.isHeader) row.dataset.joeosTableHeader = "true"; return row;')
# Render continuation header/footer variants before resolving media/font tasks.
replace_once('result.push(pageElement);', 'pageElement.joeosContinuationTemplates = [0,1].map(parity => {\n                    const holder = this.h({tagName:"div"});\n                    this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs,props,parity,false,holder);\n                    this.options.renderFooters && this.renderHeaderFooter(props.footerRefs,props,parity,false,holder);\n                    return holder;\n                });\n                result.push(pageElement);')
(out/'docx-preview.js').write_text(source,encoding='utf-8')
shutil.copyfile('node_modules/jszip/dist/jszip.min.js',out/'jszip.min.js')
for pkg in ['docx-preview','jszip']:
    for name in ['LICENSE','LICENSE.markdown']:
        p=Path('node_modules')/pkg/name
        if p.exists():shutil.copyfile(p,Path('THIRD-PARTY-LICENSES')/(pkg+'-'+name))
