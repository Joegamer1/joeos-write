import io,struct,unittest,zipfile,zlib
from package_storage import PackageParts,rewrite_package
from media_preview import wmf_bitmap_preview

class Streaming(io.BytesIO):
    def seekable(self):return False
    def seek(self,*args):raise OSError('stream')

def records(raw):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        infos=sorted(z.infolist(),key=lambda i:i.header_offset)
        return {i.filename:raw[i.header_offset:infos[n+1].header_offset if n+1<len(infos) else z.start_dir] for n,i in enumerate(infos)}

class StorageTests(unittest.TestCase):
    def test_compressed_records_descriptors_comments_and_zip64(self):
        out=Streaming()
        with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
            z.comment=b'archive comment';z.writestr('document.xml',b'old')
            with z.open('large.bin','w',force_zip64=True) as f:f.write(b'opaque'*20000)
        raw=out.getvalue();parts=PackageParts(raw)
        self.assertEqual(parts['document.xml'],b'old');self.assertNotIn('large.bin',parts.cache)
        updated=rewrite_package(raw,{'document.xml':b'new'})
        self.assertEqual(records(raw)['large.bin'],records(updated)['large.bin'])
        with zipfile.ZipFile(io.BytesIO(updated)) as z:
            self.assertIsNone(z.testzip());self.assertEqual(z.comment,b'archive comment');self.assertEqual(z.read('document.xml'),b'new')
    def test_wmf_bitmap_orientation_and_bounded_output(self):
        # Two bottom-up rows: blue/white, then red/green.
        pixels=b'\xff\0\0\xff\xff\xff\0\0'+b'\0\0\xff\0\xff\0\0\0'
        dib=struct.pack('<IiiHHIIiiII',40,2,2,1,24,0,len(pixels),0,0,0,0)
        params=struct.pack('<I8h',0xcc0020,2,2,0,0,2,2,0,0)
        record=struct.pack('<IH',(6+len(params)+len(dib)+len(pixels))//2,0xb41)+params+dib+pixels
        raw=struct.pack('<HHHIHIH',1,9,0x300,0,0,0,0)+record+struct.pack('<IH',3,0)
        png=wmf_bitmap_preview(io.BytesIO(raw),2);at=8;data=b''
        while at<len(png):
            n=struct.unpack('>I',png[at:at+4])[0]
            if png[at+4:at+8]==b'IDAT':data+=png[at+8:at+8+n]
            at+=12+n
        self.assertEqual(zlib.decompress(data),b'\0\xff\0\0\0\xff\0\0\0\0\xff\xff\xff\xff')
        with self.assertRaises(ValueError):wmf_bitmap_preview(io.BytesIO(raw[:-10]))
    def test_reject_unknown_metafile_commands(self):
        raw=struct.pack('<HHHIHIH',1,9,0x300,0,0,0,0)+struct.pack('<IH',3,0x626)
        with self.assertRaises(ValueError):wmf_bitmap_preview(io.BytesIO(raw))

if __name__=='__main__':unittest.main()
