"""Compose an app badge without altering the approved flame source."""
import os, struct
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtGui import QImage, QPainter
from PySide6.QtCore import QByteArray, QBuffer, QIODevice
app=QApplication([])
root=Path(__file__).parent/'assets'
master=(root/'M2-master.svg').read_text(encoding='utf-8')
flame=master[master.index('<path'):master.rindex('</svg>')]
svg=f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256">
<rect x="8" y="8" width="240" height="240" rx="51" fill="#174e50"/>
<g transform="translate(19 16) scale(0.41096) translate(-684 -50)">{flame}</g>
<g transform="translate(165 133) scale(.88) translate(-165 -133)">
<path d="M103 215 L111 180 L188 57 Q195 46 207 54 L219 62 Q231 70 224 81 L148 204 Z" fill="#faf7f0" stroke="#174e50" stroke-width="9" stroke-linejoin="round"/>
<path d="M118 178 L189 65 L215 82 L144 196 Z" fill="#68d6cd"/>
<path d="M189 65 L195 56 Q198 52 203 55 L219 66 Q223 69 220 74 L215 82 Z" fill="#e1fffa"/>
<path d="M107 210 L115 185 L137 200 Z" fill="#faf7f0"/>
<path d="M107 210 L111 199 L119 205 Z" fill="#292722"/>
<path d="M130 176 L193 76" fill="none" stroke="#21827e" stroke-width="5" stroke-linecap="round"/>
</g></svg>'''
(root/'write-icon.svg').write_text(svg,encoding='utf-8')
renderer=QSvgRenderer(QByteArray(svg.encode()))
blobs=[]
for size in [16,24,32,48,64,128,256]:
    im=QImage(size,size,QImage.Format_ARGB32);im.fill(0)
    p=QPainter(im);renderer.render(p);p.end()
    if size==256:im.save(str(root/'write-icon.png'))
    b=QBuffer();b.open(QIODevice.WriteOnly);im.save(b,'PNG');blobs.append((size,bytes(b.data())))
offset=6+16*len(blobs)
entries=[]
for size,data in blobs:
    entries.append(struct.pack('<BBBBHHII',size%256,size%256,0,0,1,32,len(data),offset));offset+=len(data)
(root/'write-icon.ico').write_bytes(struct.pack('<HHH',0,1,len(blobs))+b''.join(entries)+b''.join(b for _,b in blobs))
