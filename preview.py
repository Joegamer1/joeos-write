"""Deterministic screenshot of the actual Qt window, for visual review."""
import os
os.environ['QT_QPA_PLATFORM']='offscreen'
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFontDatabase
from app import Editor
app=QApplication([]); app.setStyle('Fusion')
for filename in ['segoeui.ttf','segoeuib.ttf','calibri.ttf','calibrib.ttf','calibrii.ttf']:
    QFontDatabase.addApplicationFont('C:/Windows/Fonts/'+filename)
w=Editor()
w.editor.setHtml('''<h1 style="color:#292722">A little room to think.</h1>
<p style="font-size:12pt">Welcome to your local writing workspace.</p>
<p>Start with an idea. Make it <b>bold</b>, give it <i>room to breathe</i>, and save it as a Word document when you are ready.</p>
<h2 style="color:#944018">Made for your everyday documents</h2>
<ul><li>Notes, letters, plans, and first drafts</li><li>Simple tables and inline images</li><li>Files that stay on your computer</li></ul>
<p style="color:#5c5044">JoeOS Write is a working name. The software brand can evolve later.</p>''')
w.editor.document().setModified(False)
w.show();app.processEvents();w.grab().save('preview.png')
