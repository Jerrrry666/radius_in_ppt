"""Consumer-level PPAM/OVBA format checks. Requires native/requirements-test.txt.

These tests do not execute VBA or claim successful Mac Office loading.
"""
import importlib.util
import io
import json
from pathlib import Path
import random
import re
import struct
import subprocess
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

import olefile
from oletools.olevba import VBA_Parser, decompress_stream

ROOT = Path(__file__).resolve().parent.parent
SOURCES = sorted([*(ROOT / 'native').glob('*.bas'), *(ROOT / 'native').glob('*.cls')])
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
            self.assertEqual(set(sources), {p.name for p in SOURCES})
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
        source = '\n'.join(p.read_text() for p in SOURCES)
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
        source = '\n'.join(p.read_text() for p in SOURCES)
        for unsupported in ('Declare PtrSafe', 'kernel32', 'user32', 'CreateObject(', 'Scripting.Dictionary', 'MSForms.'):
            self.assertNotIn(unsupported, source)

    def test_compression_crosses_chunk_and_sector_boundaries(self):
        for length in (1, 7, 8, 3600, 3601, 4096, 8197):
            data = bytes(i % 251 for i in range(length))
            decompressed = decompress_stream(builder.compressed(data))
            self.assertEqual(decompressed[:length], data)
            self.assertTrue(all(b == 32 for b in decompressed[length:]))

    def test_long_vba_modules_avoid_the_failing_raw_chunk_encoding(self):
        # The old literal/raw writer round-tripped in oletools but Mac Office
        # loaded the >4096-byte modules as empty, breaking Ribbon callbacks.
        for path in SOURCES:
            with self.subTest(module=path.stem):
                stream = self.ole.openstream('VBA/' + path.stem).read()
                self.assertEqual(stream[0], 1)
                position = 1
                chunks = 0
                while position < len(stream):
                    header = struct.unpack_from('<H', stream, position)[0]
                    self.assertEqual(header & 0xF000, 0xB000, 'Mac source modules must not use raw chunks')
                    position += (header & 0xFFF) + 3
                    chunks += 1
                self.assertEqual(position, len(stream))
                if len(path.read_bytes()) > 4096:
                    self.assertGreater(chunks, 1)

    def test_event_sink_is_a_class_module_in_both_project_records(self):
        project = self.ole.openstream('PROJECT').read().decode('cp1252')
        self.assertIn('Class=RadiusNativeEvents\r\n', project)
        self.assertNotIn('Module=RadiusNativeEvents', project)
        directory = decompress_stream(self.ole.openstream('VBA/dir').read())
        name = b'RadiusNativeEvents'
        start = directory.index(builder.record(0x19, name))
        end = directory.index(struct.pack('<HI', 0x2B, 0), start)
        self.assertIn(struct.pack('<HI', 0x22, 0), directory[start:end])
        self.assertIn(struct.pack('<HI', 0x28, 0), directory[start:end])
        source = decompress_stream(self.ole.openstream('VBA/RadiusNativeEvents').read()).decode('cp1252')
        # These class identity attributes are needed for Mac to load the class.
        # Merely declaring Class= and MODULETYPE 0x22 passed extraction checks
        # but left all callbacks nonfunctional in the real host.
        self.assertIn('Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"', source)
        self.assertIn('Attribute VB_TemplateDerived = False', source)
        self.assertIn('Attribute VB_Customizable = False', source)
        for path in SOURCES:
            if path.suffix != '.bas':
                continue
            start = directory.index(builder.record(0x19, path.stem.encode('cp1252')))
            end = directory.index(struct.pack('<HI', 0x2B, 0), start)
            self.assertIn(struct.pack('<HI', 0x21, 0), directory[start:end])

    def test_native_ribbon_shows_the_active_release_version(self):
        ui = ET.fromstring(self.archive.read('customUI/customUI.xml'))
        version = next(e for e in ui.iter() if e.get('id') == 'NativeVersion')
        package = json.loads((ROOT / 'package.json').read_text())
        self.assertEqual(version.get('label'), 'v' + package['version'])

    def test_every_action_has_an_embedded_png_relationship(self):
        ui = ET.fromstring(self.archive.read('customUI/customUI.xml'))
        relationships = {r.get('Id'): r for r in ET.fromstring(self.archive.read('customUI/_rels/customUI.xml.rels'))}
        types = ET.fromstring(self.archive.read('[Content_Types].xml'))
        self.assertTrue(any(e.get('Extension') == 'png' and e.get('ContentType') == 'image/png' for e in types))
        images = set()
        for control in ui.iter():
            if control.tag.split('}')[-1] not in ('button', 'toggleButton'):
                continue
            name = control.get('image')
            self.assertIsNotNone(name, control.get('id'))
            images.add(name)
            relationship = relationships[name]
            self.assertTrue(relationship.get('Type').endswith('/image'))
            part = 'customUI/' + relationship.get('Target')
            payload = self.archive.read(part)
            self.assertEqual(payload, (ROOT / 'native' / 'icons' / (name + '.png')).read_bytes())
            self.assertEqual(payload[:8], b'\x89PNG\r\n\x1a\n')
            self.assertEqual(struct.unpack('>II', payload[16:24]), (32, 32))
        self.assertEqual(images, set(relationships))

    def test_missing_or_invalid_image_fails_before_packaging(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = Path(directory)
            (fixture / 'Module.bas').write_text('Attribute VB_Name = "Module"\nOption Explicit\n')
            (fixture / 'customUI.xml').write_text('<customUI><button id="b" image="missing"/></customUI>')
            with self.assertRaisesRegex(ValueError, 'Missing ribbon image'):
                builder.build(fixture / 'bad.ppam', fixture)
            (fixture / 'icons').mkdir()
            (fixture / 'icons/missing.png').write_bytes(b'bad png')
            with self.assertRaisesRegex(ValueError, 'Invalid PNG'):
                builder.build(fixture / 'bad.ppam', fixture)

    def test_token_compression_overlapping_copies_and_offset_bit_boundaries(self):
        rng = random.Random(20261007)
        cases = [b'A' * 10000, b'abc' * 3500, b'Option Explicit\r\n' * 900]
        for boundary in (15, 16, 17, 31, 32, 33, 63, 64, 65, 255, 256, 257,
                         511, 512, 513, 1023, 1024, 1025, 2047, 2048, 2049):
            prefix = rng.randbytes(boundary)
            cases.append(prefix + prefix[:100] + b'Z' * 64)
        for data in cases:
            with self.subTest(length=len(data)):
                self.assertEqual(decompress_stream(builder.compressed(data)), data)

    def test_incompressible_full_chunks_and_short_tail_round_trip(self):
        data = random.Random(2).randbytes(8197)
        stream = builder.compressed(data)
        self.assertEqual(struct.unpack_from('<H', stream, 1)[0], 0x3FFF)
        self.assertEqual(decompress_stream(stream), data)

    def test_distribution_contains_only_the_embedded_plugin_and_installation_files(self):
        output = builder.distribution(self.output)
        first = output.read_bytes()
        with zipfile.ZipFile(output) as archive:
            prefix = 'RadiusInPptNative-mac/'
            self.assertEqual(set(archive.namelist()), {
                prefix + 'RadiusInPptNative.ppam',
                prefix + 'Install-RadiusInPptNative.command',
                prefix + 'INSTALL.txt',
            })
            self.assertEqual(archive.read(prefix + 'RadiusInPptNative.ppam'), self.output.read_bytes())
            mode = archive.getinfo(prefix + 'Install-RadiusInPptNative.command').external_attr >> 16
            self.assertEqual(mode & 0o777, 0o755)
        self.assertEqual(builder.distribution(self.output).read_bytes(), first)

    def test_install_helper_handles_spaces_reinstall_and_preserves_previous_package(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            package = directory / 'Extracted package'
            package.mkdir()
            helper = package / 'Install-RadiusInPptNative.command'
            helper.write_bytes((ROOT / 'native' / helper.name).read_bytes())
            payload = package / 'RadiusInPptNative.ppam'
            payload.write_bytes(self.output.read_bytes())
            destination = directory / 'Mac Library' / '原生插件'
            command = ['bash', str(helper), '--destination', str(destination), '--no-reveal']
            target = destination / payload.name
            previous = destination / 'RadiusInPptNative.previous.ppam'
            subprocess.run(command, check=True, capture_output=True)
            self.assertEqual(target.read_bytes(), payload.read_bytes())
            self.assertFalse(previous.exists())
            target_stat = target.stat()
            subprocess.run(command, check=True, capture_output=True)
            self.assertEqual(target.stat().st_mtime_ns, target_stat.st_mtime_ns)
            self.assertFalse(previous.exists())
            old = target.read_bytes()
            with zipfile.ZipFile(payload, 'a') as archive:
                archive.writestr('docProps/test-update.txt', 'new package')
            subprocess.run(command, check=True, capture_output=True)
            self.assertEqual(target.read_bytes(), payload.read_bytes())
            self.assertEqual(previous.read_bytes(), old)
            self.assertEqual(target.stat().st_mode & 0o777, 0o644)
            self.assertFalse(list(destination.glob('.radius-install.*')))

    def test_install_helper_rejects_missing_or_corrupt_payload_before_writing(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            helper = directory / 'Install.command'
            helper.write_bytes((ROOT / 'native' / 'Install-RadiusInPptNative.command').read_bytes())
            destination = directory / 'Untouched'
            command = ['bash', str(helper), '--destination', str(destination), '--no-reveal']
            for payload in (None, b'not a PowerPoint ZIP'):
                if payload is not None:
                    (directory / 'RadiusInPptNative.ppam').write_bytes(payload)
                result = subprocess.run(command, capture_output=True)
                self.assertNotEqual(result.returncode, 0)
                self.assertFalse(destination.exists())

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
