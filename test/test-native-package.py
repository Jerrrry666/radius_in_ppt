"""Consumer-level PPAM/OVBA format checks. Requires native/requirements-test.txt.

These tests do not execute VBA or claim successful Mac Office loading.
"""
import importlib.util
import io
from pathlib import Path
import re
import struct
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

import olefile
from oletools.olevba import VBA_Parser, decompress_stream

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('build_native', ROOT / 'tools' / 'build-native.py')
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)


class NativePackageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.output = builder.build(Path(cls.temp.name) / 'radius.ppam')
        cls.archive = zipfile.ZipFile(cls.output)
        cls.ole = olefile.OleFileIO(io.BytesIO(cls.archive.read('ppt/vbaProject.bin')))

    @classmethod
    def tearDownClass(cls):
        cls.ole.close()
        cls.archive.close()
        cls.temp.cleanup()

    def test_macro_enabled_addin_type_and_relationship_targets(self):
        types = ET.fromstring(self.archive.read('[Content_Types].xml'))
        self.assertTrue(any(e.get('ContentType') == 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml' for e in types))
        for name in self.archive.namelist():
            if not name.endswith('.rels'):
                continue
            folder = '' if name == '_rels/.rels' else name.split('/_rels/')[0] + '/'
            for relationship in ET.fromstring(self.archive.read(name)):
                self.assertIn(folder + relationship.get('Target'), self.archive.namelist())

    def test_oletools_extracts_all_original_sources_without_performance_cache(self):
        parser = VBA_Parser(str(self.output))
        try:
            self.assertTrue(parser.detect_vba_macros())
            sources = {name: source for _, _, name, source in parser.extract_macros()}
            self.assertEqual(set(sources), {p.name for p in (ROOT / 'native').glob('*.bas')})
            for name, source in sources.items():
                original = (ROOT / 'native' / name).read_text().replace('\r\n', '\n')
                self.assertEqual(source.replace('\r\n', '\n').rstrip(' '), original)
        finally:
            parser.close()
        self.assertEqual(self.ole.openstream('VBA/_VBA_PROJECT').read(), struct.pack('<HHBH', 0x61CC, 0xFFFF, 0, 0))

    def test_mac_platform_and_mac_portable_library_references(self):
        directory = decompress_stream(self.ole.openstream('VBA/dir').read())
        self.assertEqual(struct.unpack_from('<HII', directory), (1, 4, 2))
        self.assertIn(b'PowerPoint', directory)
        self.assertIn(b'{91493440-5A91-11CF-8700-00AA0060263B}', directory)
        self.assertNotIn(b'C:\\', directory)

    def test_ribbon_callbacks_match_extracted_modules_and_have_unique_ids(self):
        source = '\n'.join(p.read_text() for p in (ROOT / 'native').glob('*.bas'))
        callbacks = set(re.findall(r'Public Sub (\w+)\(', source))
        ids = set()
        for element in ET.fromstring(self.archive.read('customUI/customUI.xml')).iter():
            if element.get('id'):
                self.assertNotIn(element.get('id'), ids)
                ids.add(element.get('id'))
            for key, value in element.attrib.items():
                if key.startswith('on') or key.startswith('get'):
                    self.assertIn(value, callbacks)
        self.assertIn('NativeRadiusValue', ids)

    def test_public_core_is_source_only_and_no_windows_or_com_factory_dependencies(self):
        source = '\n'.join(p.read_text() for p in (ROOT / 'native').glob('*.bas'))
        for unsupported in ('Declare PtrSafe', 'kernel32', 'user32', 'CreateObject(', 'Scripting.Dictionary', 'MSForms.'):
            self.assertNotIn(unsupported, source)

    def test_compression_crosses_chunk_and_sector_boundaries(self):
        for length in (1, 7, 8, 3600, 3601, 4096, 8197):
            data = bytes(i % 251 for i in range(length))
            decompressed = decompress_stream(builder.compressed(data))
            self.assertEqual(decompressed[:length], data)
            self.assertTrue(all(b == 32 for b in decompressed[length:]))

    def test_build_is_reproducible_and_missing_callback_fails(self):
        other = builder.build(Path(self.temp.name) / 'second.ppam')
        self.assertEqual(other.read_bytes(), self.output.read_bytes())
        fixture = Path(self.temp.name) / 'bad-native'
        fixture.mkdir()
        (fixture / 'Module.bas').write_text('Attribute VB_Name = "Module"\nOption Explicit\n')
        (fixture / 'customUI.xml').write_text('<customUI onLoad="Missing"/>')
        with self.assertRaisesRegex(ValueError, 'Missing callback'):
            builder.build(Path(self.temp.name) / 'bad.ppam', fixture)


if __name__ == '__main__':
    unittest.main(verbosity=2)
