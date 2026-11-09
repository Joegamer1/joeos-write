"""Load installed fonts for local display; never change or embed DOCX fonts."""
import base64
import io
import os
from pathlib import Path
from fontTools.ttLib import TTCollection, TTFont
from PySide6.QtGui import QRawFont, QFont

# Fallbacks only when an exact installed outline family is absent. Earlier
# versions may have different metrics; report the substitution to the user.
COMPATIBILITY_FONTS = {'bierstadt': 'Aptos', 'courier': 'Courier New',
                       'times new roman (headings cs)': 'Times New Roman'}

class LocalFonts:
    def __init__(self, roots=None):
        self.cache = {}
        self.index = None
        self.resolutions = []
        self.roots = roots if roots is not None else [
            Path.home()/'.joeos/fonts',
            Path(os.environ.get('LOCALAPPDATA', Path.home()/'AppData/Local'))/'Microsoft/Windows/Fonts',
            Path(os.environ.get('WINDIR', 'C:/Windows'))/'Fonts']

    def _index(self):
        if self.index is not None:
            return
        self.index = []
        for root in self.roots:
            try:
                paths = sorted(Path(root).iterdir())
            except OSError:
                continue
            for path in paths:
                if path.suffix.lower() not in ('.ttf', '.otf', '.ttc', '.otc'):
                    continue
                try:
                    if path.stat().st_size > 64*1024*1024:
                        continue
                    if path.suffix.lower() in ('.ttc', '.otc'):
                        collection = TTCollection(path, lazy=True)
                        try:
                            for number, font in enumerate(collection.fonts):
                                names = font['name']
                                family = names.getDebugName(16) or names.getDebugName(1)
                                weight = font['OS/2'].usWeightClass
                                style = 'italic' if font['head'].macStyle & 2 else 'normal'
                                if family:
                                    self.index.append((family, weight, style, path, number))
                        finally:
                            collection.close()
                    else:
                        font = QRawFont(str(path), 16)
                        if font.isValid():
                            style = 'normal' if font.style() == QFont.StyleNormal else 'italic'
                            self.index.append((font.familyName(), font.weight(), style, path, None))
                except Exception:
                    # A damaged or inaccessible local font must not block Open.
                    continue

    def _data(self, path, number):
        key = (path, number)
        if key not in self.cache:
            if number is None:
                raw = path.read_bytes()
            else:
                # Serialize the selected original TTC face in memory. Retain
                # glyphs and shaping tables; do not alter the installed file.
                with TTFont(path, fontNumber=number, recalcTimestamp=False) as font:
                    output = io.BytesIO()
                    font.save(output)
                    raw = output.getvalue()
            self.cache[key] = base64.b64encode(raw).decode('ascii')
        return self.cache[key]

    def faces(self, families):
        self._index()
        available = {f[0].casefold() for f in self.index}
        result, seen, self.resolutions = [], set(), []
        total = 0
        for requested in sorted(set(families), key=str.casefold):
            target = requested if requested.casefold() in available else COMPATIBILITY_FONTS.get(requested.casefold(), requested)
            loaded = False
            for family, weight, style, path, number in self.index:
                identity = (requested.casefold(), weight, style)
                if family.casefold() != target.casefold() or identity in seen:
                    continue
                try:
                    data = self._data(path, number)
                    if total + len(data) > 48*1024*1024:
                        continue
                    total += len(data)
                    seen.add(identity)
                    result.append({'family': requested, 'sourceFamily': family,
                                   'weight': str(weight), 'style': style, 'data': data})
                    loaded = True
                except Exception:
                    continue
            self.resolutions.append({'requested': requested, 'sourceFamily': target if loaded else None,
                                     'compatibility': loaded and target.casefold() != requested.casefold()})
        return result
