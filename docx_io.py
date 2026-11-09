"""Deliberately bounded DOCX <-> Qt rich-text conversion. No network IO."""
import io
import os
import tempfile
import zipfile
import math
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from PySide6.QtCore import Qt, QUrl, QByteArray, QBuffer, QIODevice
from PySide6.QtGui import (QTextDocument, QTextCursor, QTextCharFormat, QTextBlockFormat,
                          QTextListFormat, QTextTableFormat, QTextImageFormat, QImage, QFont)


def font_points(font):
    """Read half-point values tolerantly: some generators emit decimal w:sz."""
    if font is None:
        return None
    props = font._element.rPr
    size = props.find(qn('w:sz')) if props is not None else None
    if size is None:
        return None
    try:
        value = float(size.get(qn('w:val'))) / 2
        return value if math.isfinite(value) and 0 < value <= 1638 else None
    except (TypeError, ValueError):
        return None


def load_docx(path):
    with zipfile.ZipFile(path) as z:
        if sum(i.file_size for i in z.infolist()) > 100_000_000:
            raise ValueError('Document is too large to safely open (100 MB expanded limit).')
    src = Document(path)
    out = QTextDocument()
    out.setDefaultFont(QFont('Calibri', 11))
    cursor = QTextCursor(out)
    warnings = set()
    xml = src.element.xml
    for token, label in [('sectPr','Page layout'), ('drawing','Image placement'), ('fldChar','Fields'),
                         ('sdt','Content controls'), ('ins','Tracked changes'), ('del','Tracked changes'),
                         ('footnoteReference','Footnotes'), ('commentRangeStart','Comments')]:
        if '<w:' + token in xml:
            warnings.add(label)
    if any(s.header.paragraphs[0].text or s.footer.paragraphs[0].text for s in src.sections):
        warnings.add('Headers and footers')
    image_index = 0

    def paragraph(p, c):
        nonlocal image_index
        bf = QTextBlockFormat()
        align = int(p.alignment) if p.alignment is not None else 0
        bf.setAlignment({0:Qt.AlignLeft,1:Qt.AlignHCenter,2:Qt.AlignRight,3:Qt.AlignJustify}.get(align, Qt.AlignLeft))
        name = p.style.name if p.style else ''
        if name.startswith('Heading ') and name[-1:].isdigit():
            bf.setHeadingLevel(int(name[-1]))
        c.setBlockFormat(bf)
        if 'List' in name or p._p.find('.//' + qn('w:numPr')) is not None:
            lf = QTextListFormat()
            lf.setStyle(QTextListFormat.ListDecimal if 'Number' in name else QTextListFormat.ListDisc)
            c.createList(lf)
            warnings.add('List numbering and nesting')
        # Include visible hyperlink text; hyperlink destinations are not preserved.
        from docx.text.run import Run
        for child in p._p:
            nodes = [child] if child.tag == qn('w:r') else list(child.iter(qn('w:r')))
            if child.tag == qn('w:hyperlink'):
                warnings.add('Hyperlink destinations')
            for node in nodes:
                r = Run(node, p)
                cf = QTextCharFormat()
                stylefont = p.style.font if p.style else None
                cf.setFontWeight(QFont.Bold if (r.bold if r.bold is not None else (stylefont.bold if stylefont else False)) else QFont.Normal)
                cf.setFontItalic(bool(r.italic if r.italic is not None else (stylefont.italic if stylefont else False)))
                cf.setFontUnderline(bool(r.underline))
                cf.setFontStrikeOut(bool(r.font.strike))
                cf.setFontFamilies([r.font.name or (stylefont.name if stylefont else None) or 'Calibri'])
                size = font_points(r.font) or font_points(stylefont)
                cf.setFontPointSize(size if size else (max(12, 26-bf.headingLevel()*4) if bf.headingLevel() else 11))
                if r.font.color.rgb:
                    from PySide6.QtGui import QColor
                    cf.setForeground(QColor('#'+str(r.font.color.rgb)))
                c.insertText(r.text, cf)
                for blip in node.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                    rid = blip.get(qn('r:embed'))
                    if rid and rid in src.part.related_parts:
                        data = src.part.related_parts[rid].blob
                        im = QImage.fromData(data)
                        if not im.isNull():
                            key = 'embedded-image-'+str(image_index)
                            image_index += 1
                            out.addResource(QTextDocument.ImageResource, QUrl(key), im)
                            fmt = QTextImageFormat(); fmt.setName(key)
                            extent = node.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
                            fmt.setWidth(int(extent.get('cx')) / 9525 if extent is not None else min(im.width(),600))
                            fmt.setHeight(int(extent.get('cy')) / 9525 if extent is not None else im.height()*fmt.width()/im.width())
                            c.insertImage(fmt)
                        else:
                            warnings.add('Unsupported image format')

    from docx.text.paragraph import Paragraph
    from docx.table import Table
    first = True
    for elem in src.element.body:
        if elem.tag == qn('w:p'):
            if not first: cursor.insertBlock()
            paragraph(Paragraph(elem, src), cursor)
            first = False
        elif elem.tag == qn('w:tbl'):
            t = Table(elem, src)
            tf = QTextTableFormat(); tf.setBorder(1); tf.setCellPadding(6)
            qt = cursor.insertTable(len(t.rows), len(t.columns), tf)
            for i, row in enumerate(t.rows):
                for j, cell in enumerate(row.cells):
                    cc = qt.cellAt(i,j).firstCursorPosition()
                    for k,p in enumerate(cell.paragraphs):
                        if k: cc.insertBlock()
                        paragraph(p,cc)
            cursor = qt.lastCursorPosition(); cursor.movePosition(QTextCursor.NextBlock)
            first = True
            warnings.add('Table widths, merged cells and nested tables')
    out.setModified(False)
    return out, sorted(warnings)


def save_docx(document, path):
    """Write a complete package to a sibling temp file, then replace atomically."""
    dst = Document()
    dst.styles['Normal'].font.name = 'Calibri'
    dst.styles['Normal'].font.size = Pt(11)

    def block_to_paragraph(block, parent):
        p = parent.add_paragraph()
        bf = block.blockFormat()
        if bf.headingLevel(): p.style = 'Heading '+str(min(9,bf.headingLevel()))
        if block.textList():
            decimal = block.textList().format().style() == QTextListFormat.ListDecimal
            p.style = 'List Number' if decimal else 'List Bullet'
        a = bf.alignment()
        p.alignment = 3 if a & Qt.AlignJustify else 2 if a & Qt.AlignRight else 1 if a & Qt.AlignHCenter else 0
        it = block.begin()
        while not it.atEnd():
            fragment = it.fragment(); it += 1
            if not fragment.isValid(): continue
            fmt = fragment.charFormat()
            if fmt.isImageFormat():
                imgfmt = fmt.toImageFormat()
                data = document.resource(QTextDocument.ImageResource, QUrl(imgfmt.name()))
                im = data if isinstance(data, QImage) else QImage(imgfmt.name())
                if im.isNull(): raise ValueError('An image could not be saved. Original file was not changed.')
                buf = QBuffer(); buf.open(QIODevice.WriteOnly); im.save(buf,'PNG')
                p.add_run().add_picture(io.BytesIO(bytes(buf.data())), width=Inches((imgfmt.width() or im.width())/96))
                continue
            r = p.add_run(fragment.text().replace('\u2028','\n'))
            r.bold = fmt.fontWeight() >= QFont.Bold
            r.italic = fmt.fontItalic(); r.underline = fmt.fontUnderline()
            r.font.strike = fmt.fontStrikeOut()
            families = fmt.fontFamilies()
            r.font.name = families[0] if families else 'Calibri'
            r.font.size = Pt(fmt.fontPointSize() or 11)
            if fmt.foreground().style() != Qt.NoBrush:
                r.font.color.rgb = RGBColor.from_string(fmt.foreground().color().name()[1:])
        return p

    def walk(frame, parent):
        it = frame.begin()
        while not it.atEnd():
            child = it.currentFrame()
            if child is not None:
                from PySide6.QtGui import QTextTable
                if isinstance(child,QTextTable):
                    table = parent.add_table(rows=child.rows(), cols=child.columns())
                    table.style = 'Table Grid'
                    for i in range(child.rows()):
                        for j in range(child.columns()):
                            cell = child.cellAt(i,j); target = table.cell(i,j)
                            b = cell.firstCursorPosition().block()
                            end = cell.lastCursorPosition().position()
                            while b.isValid() and b.position() <= end:
                                block_to_paragraph(b,target); b = b.next()
                            empty = target.paragraphs[0]._element
                            empty.getparent().remove(empty)
                else: walk(child,parent)
            elif it.currentBlock().isValid(): block_to_paragraph(it.currentBlock(),parent)
            it += 1
    walk(document.rootFrame(),dst)
    path = Path(path)
    fd, tmp = tempfile.mkstemp(prefix='.joeos-',suffix='.docx',dir=path.parent)
    os.close(fd)
    try:
        dst.save(tmp)
        Document(tmp)  # Reopen the generated package before replacement.
        os.replace(tmp,path)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)
