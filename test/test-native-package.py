"""Consumer-level PPAM/OVBA format checks. Requires native/requirements-test.txt.

These tests do not execute VBA or claim successful Mac Office loading.
"""
import importlib.util
import hashlib
import io
import json
import os
from pathlib import Path
import pty
import random
import re
import select
import struct
import subprocess
import tempfile
import time
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
fixture_spec = importlib.util.spec_from_file_location('native_host_fixture', ROOT / 'test/native-host-fixture.py')
host_fixture = importlib.util.module_from_spec(fixture_spec)
fixture_spec.loader.exec_module(host_fixture)


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
            if control.tag.split('}')[-1] not in ('button', 'toggleButton', 'menu', 'dynamicMenu'):
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

    def test_relation_controls_and_localized_dynamic_labels_are_packaged(self):
        ui = ET.fromstring(self.archive.read('customUI/customUI.xml'))
        controls = {e.get('id'): e for e in ui.iter() if e.get('id')}
        expected = {'NativeRelParent', 'NativeRelBind', 'NativeRelCancel', 'NativeRelView',
                    'NativeRelPreview', 'NativeRelDetach', 'NativeRelRemoveAll',
                    'NativeRelStatus', 'NativeRelInfo'}
        self.assertTrue(expected <= controls.keys())
        self.assertEqual(controls['NativeRelView'].get('getContent'), 'NativeRelGetMenu')
        self.assertEqual(controls['NativeRelPreview'].get('onAction'), 'NativeRelPreview')
        self.assertIn('副本', controls['NativeRelPreview'].get('tag'))
        for source in SOURCES:
            source.read_text().encode('cp1252')

    def test_layout_configuration_and_modes_are_exposed_in_ribbon(self):
        ui = ET.fromstring(self.archive.read('customUI/customUI.xml'))
        controls = {e.get('id'): e for e in ui.iter() if e.get('id')}
        for name in ('Rows', 'Columns', 'Padding', 'Gap'):
            field = controls['NativeLayout' + name]
            self.assertEqual(field.tag.split('}')[-1], 'editBox')
            self.assertEqual(field.get('getText'), 'NativeLayoutGetText')
            self.assertEqual(field.get('onChange'), 'NativeLayoutParameter')
            self.assertEqual(field.get('showLabel'), 'false')
            steps = list(controls['NativeLayout' + name + 'Steps'])
            self.assertEqual([e.get('tag') for e in steps],
                             [field.get('tag') + '|up', field.get('tag') + '|down'])
            for step, image in zip(steps, ('stepUp', 'stepDown')):
                self.assertEqual(step.get('showLabel'), 'false')
                self.assertEqual(step.get('image'), image)
                self.assertEqual(step.get('getEnabled'), 'NativeLayoutGetStepEnabled')
                self.assertEqual(step.get('onAction'), 'NativeLayoutStep')
        modes = list(controls['NativeLayoutMode'])
        self.assertEqual([e.get('id') for e in modes],
                         ['NativeLayoutSame', 'NativeLayoutSubtract', 'NativeLayoutOff'])
        self.assertEqual(controls['NativeLayoutAuto'].get('getPressed'), 'NativeLayoutAutoPressed')
        self.assertEqual(controls['NativeLayoutApply'].get('onAction'), 'NativeLayoutApply')

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

    def test_install_helper_handles_spaces_reinstall_and_updates_without_backup(self):
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
            subprocess.run(command, check=True, capture_output=True)
            self.assertEqual(target.read_bytes(), payload.read_bytes())
            self.assertEqual({path.name for path in destination.iterdir()}, {payload.name})
            target_stat = target.stat()
            subprocess.run(command, check=True, capture_output=True)
            self.assertEqual(target.stat().st_mtime_ns, target_stat.st_mtime_ns)
            self.assertEqual({path.name for path in destination.iterdir()}, {payload.name})
            old = target.read_bytes()
            with zipfile.ZipFile(payload, 'a') as archive:
                archive.writestr('docProps/test-update.txt', 'new package')
            subprocess.run(command, check=True, capture_output=True)
            self.assertEqual(target.read_bytes(), payload.read_bytes())
            self.assertNotEqual(target.read_bytes(), old)
            self.assertEqual(target.stat().st_mode & 0o777, 0o644)
            self.assertEqual({path.name for path in destination.iterdir()}, {payload.name})

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

    def test_install_helper_failed_update_preserves_current_package_and_cleans_staging(self):
        # Inject filesystem failures inside this subprocess only. The current
        # install must remain intact before and after the new copy completes,
        # without a test hook in the installer or a separate recovery backup.
        fault_script = r'''
fail_phase="$1"
shift
fail_copy_source="$1"
shift
fail_replace_target="$1"
shift
function /bin/cp {
  if [ "$fail_phase" = 'copy' ] && [ "$1" = "$fail_copy_source" ]; then
    printf 'interrupted copy' > "${@: -1}"
    echo '[test] injected filesystem failure' >&2
    return 73
  fi
  command /bin/cp "$@"
}
function /bin/mv {
  if [ "$fail_phase" = 'replace' ] && [ "${@: -1}" = "$fail_replace_target" ]; then
    echo '[test] injected filesystem failure' >&2
    return 73
  fi
  command /bin/mv "$@"
}
source "$0" "$@"
'''
        for phase in ('copy', 'replace'):
            with self.subTest(phase=phase), tempfile.TemporaryDirectory() as directory:
                directory = Path(directory)
                helper = directory / 'Install.command'
                helper.write_bytes((ROOT / 'native/Install-RadiusInPptNative.command').read_bytes())
                payload = directory / 'RadiusInPptNative.ppam'
                payload.write_bytes(self.output.read_bytes())
                with zipfile.ZipFile(payload, 'a') as archive:
                    archive.writestr('docProps/update.txt', 'new package')
                destination = directory / 'installed'
                destination.mkdir()
                target = destination / payload.name
                target.write_bytes(self.output.read_bytes())
                before = {path.name: (hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns)
                          for path in destination.iterdir()}
                result = subprocess.run(['bash', '-c', fault_script, str(helper), phase, str(payload), str(target),
                                         '--destination', str(destination), '--no-reveal'],
                                        capture_output=True, text=True)
                self.assertEqual(result.returncode, 73, result.stderr)
                self.assertIn('[test] injected filesystem failure', result.stderr)
                after = {path.name: (hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns)
                         for path in destination.iterdir()}
                self.assertEqual(after, before)

    def test_install_helper_rejects_nonregular_package_target(self):
        for kind in ('directory', 'fifo'):
            with self.subTest(kind=kind), tempfile.TemporaryDirectory() as directory:
                directory = Path(directory)
                helper = directory / 'Install.command'
                helper.write_bytes((ROOT / 'native/Install-RadiusInPptNative.command').read_bytes())
                payload = directory / 'RadiusInPptNative.ppam'
                payload.write_bytes(self.output.read_bytes())
                destination = directory / 'installed'
                destination.mkdir()
                invalid = destination / payload.name
                if kind == 'directory':
                    invalid.mkdir()
                else:
                    os.mkfifo(invalid)
                invalid_stat = invalid.stat()
                other_name = 'Unrelated.txt'
                other = destination / other_name
                other.write_bytes(b'existing package')
                before = other.read_bytes()
                result = subprocess.run(['bash', str(helper), '--destination', str(destination), '--no-reveal'],
                                        capture_output=True, text=True)
                self.assertNotEqual(result.returncode, 0)
                self.assertEqual(other.read_bytes(), before)
                self.assertEqual(invalid.stat(), invalid_stat)
                if kind == 'directory':
                    self.assertEqual(list(invalid.iterdir()), [])
                self.assertEqual({path.name for path in destination.iterdir()}, {payload.name, other_name})

    def test_build_is_reproducible_and_missing_callback_fails(self):
        other = builder.build(Path(self.temp.name) / 'second.ppam')
        self.assertEqual(other.read_bytes(), self.output.read_bytes())
        fixture = Path(self.temp.name) / 'bad-native'
        fixture.mkdir()
        (fixture / 'Module.bas').write_text('Attribute VB_Name = "Module"\nOption Explicit\n')
        (fixture / 'customUI.xml').write_text('<customUI onLoad="Missing"/>')
        with self.assertRaisesRegex(ValueError, 'Missing callback'):
            builder.build(Path(self.temp.name) / 'bad.ppam', fixture)


class LocalTestEntryTests(unittest.TestCase):
    @staticmethod
    def prepare_entry(directory, status):
        project = Path(directory) / 'Project with spaces' / '原生插件'
        tools = project / 'tools'
        tools.mkdir(parents=True)
        entry = tools / 'Run-Tests.command'
        entry.write_bytes((ROOT / 'tools/Run-Tests.command').read_bytes())
        entry.chmod(0o755)
        (project / 'package.json').write_bytes((ROOT / 'package.json').read_bytes())
        binaries = Path(directory) / 'fake-bin'
        binaries.mkdir()
        venv_bin = project / '.venv-native/bin'
        venv_bin.mkdir(parents=True)
        for executable in (binaries / 'python3', binaries / 'node', venv_bin / 'python3'):
            executable.write_text('#!/bin/bash\nexit 0\n')
            executable.chmod(0o755)
        npm = binaries / 'npm'
        npm.write_text(r'''#!/bin/bash
printf '%s\n' "$PWD" "$*" "$(command -v python3)" > "$RADIUS_TEST_CAPTURE"
echo '[test] local suite result'
exit "$RADIUS_TEST_EXIT"
''')
        npm.chmod(0o755)
        capture = Path(directory) / 'entry-capture.txt'
        environment = dict(os.environ, PATH=str(binaries) + os.pathsep + os.defpath,
                           RADIUS_TEST_CAPTURE=str(capture), RADIUS_TEST_EXIT=str(status))
        return entry, project, capture, environment

    def test_local_entry_handles_spaces_venv_and_suite_exit_status(self):
        for status in (0, 41):
            with self.subTest(status=status), tempfile.TemporaryDirectory() as directory:
                entry, project, capture, environment = self.prepare_entry(directory, status)
                result = subprocess.run([str(entry), '--no-pause'], cwd=directory, env=environment,
                                        capture_output=True, text=True, timeout=5)
                self.assertEqual(result.returncode, status, result.stderr)
                self.assertEqual(capture.read_text().splitlines(),
                                 [str(project), 'run test:all', str(project / '.venv-native/bin/python3')])
                self.assertIn('全部通过' if status == 0 else '退出码 41', result.stdout + result.stderr)

    def test_local_entry_keeps_failed_interactive_result_visible(self):
        with tempfile.TemporaryDirectory() as directory:
            entry, _, _, environment = self.prepare_entry(directory, 41)
            master, slave = pty.openpty()
            process = subprocess.Popen([str(entry)], cwd=directory, env=environment,
                                       stdin=slave, stdout=slave, stderr=slave)
            os.close(slave)
            try:
                output = b''
                deadline = time.monotonic() + 5
                while '按回车结束检查'.encode() not in output:
                    remaining = deadline - time.monotonic()
                    self.assertGreater(remaining, 0, output.decode(errors='replace'))
                    readable, _, _ = select.select([master], [], [], remaining)
                    self.assertTrue(readable, output.decode(errors='replace'))
                    output += os.read(master, 8192)
                self.assertIsNone(process.poll(), 'Failed result closed before confirmation')
                self.assertIn('退出码 41'.encode(), output)
                os.write(master, b'\n')
                self.assertEqual(process.wait(timeout=5), 41)
            finally:
                if process.poll() is None:
                    process.kill()
                    process.wait(timeout=5)
                os.close(master)


class NativeHostFixtureTests(unittest.TestCase):
    @staticmethod
    def write_inspection_fixture(path, second_name='Second', nested=False):
        """Minimal saved OOXML, independent of PowerPoint and python-pptx."""
        p, a = host_fixture.NS['p'], host_fixture.NS['a']
        root = ET.Element('{%s}sld' % p)
        tree = ET.SubElement(ET.SubElement(root, '{%s}cSld' % p), '{%s}spTree' % p)

        def add_shape(parent, name, identifier, group=False):
            shape = ET.SubElement(parent, '{%s}%s' % (p, 'grpSp' if group else 'sp'))
            nv = ET.SubElement(shape, '{%s}%s' % (p, 'nvGrpSpPr' if group else 'nvSpPr'))
            ET.SubElement(nv, '{%s}cNvPr' % p, id=str(identifier), name=name)
            ET.SubElement(nv, '{%s}nvPr' % p)
            properties = ET.SubElement(shape, '{%s}%s' % (p, 'grpSpPr' if group else 'spPr'))
            transform = ET.SubElement(properties, '{%s}xfrm' % a)
            ET.SubElement(transform, '{%s}off' % a, x='0', y='0')
            ET.SubElement(transform, '{%s}ext' % a, cx='360000', cy='360000')
            if group:
                ET.SubElement(transform, '{%s}chOff' % a, x='0', y='0')
                ET.SubElement(transform, '{%s}chExt' % a, cx='360000', cy='360000')
            return shape

        add_shape(tree, 'First', 2)
        container = add_shape(tree, 'Group', 3, True) if nested else tree
        add_shape(container, second_name, 4)
        with zipfile.ZipFile(path, 'w') as archive:
            archive.writestr('ppt/slides/slide1.xml', ET.tostring(root))
            archive.writestr('ppt/slides/_rels/slide1.xml.rels',
                             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')

    def test_inspect_preserves_unique_shape_names_ids_and_group_hierarchy(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'unique.pptx'
            self.write_inspection_fixture(path, nested=True)
            shapes = host_fixture.inspect(path, 1)['1']
            self.assertEqual(set(shapes), {'First', 'Group', 'Second'})
            self.assertEqual(shapes['First']['id'], '2')
            self.assertEqual(shapes['Second']['id'], '4')
            self.assertEqual(shapes['Second']['parent'], 'Group')
            self.assertEqual(shapes['Second']['boxCm'], [0, 0, 1, 1])

    def test_inspect_rejects_duplicate_names_in_top_level_and_nested_shapes(self):
        for nested in (False, True):
            with self.subTest(nested=nested), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / 'duplicate.pptx'
                self.write_inspection_fixture(path, second_name='First', nested=nested)
                with self.assertRaisesRegex(ValueError, "Duplicate shape name 'First' on slide 1.*IDs 2 and 4"):
                    host_fixture.inspect(path, 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
