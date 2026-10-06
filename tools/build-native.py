#!/usr/bin/env python3
"""Build a source-only, Mac-targeted PPAM without launching Office.

Uses MS-CFB/MS-OVBA and OPC. Office must compile the source on first load;
format validation is not a claim of successful Mac VBA compilation.
"""
import argparse
import math
from pathlib import Path
import re
import shutil
import struct
import xml.etree.ElementTree as ET
import zipfile

FREE, END, FAT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
ROOT = Path(__file__).resolve().parent.parent


def compressed(data):
    """MS-OVBA tokens, including overlapping copies and per-chunk windows.

    Avoid the raw-chunk encoding used in the failing long-module packages.
    VBA source uses copy tokens, preserving every byte across chunk boundaries.
    Incompressible binary input still uses the standard 4096-byte raw chunk.
    """
    out = bytearray(b'\x01')
    for offset in range(0, len(data), 4096):
        block = data[offset:offset + 4096]
        body = bytearray()
        position = 0
        while position < len(block):
            flag_index = len(body)
            body.append(0)
            for bit in range(8):
                if position == len(block):
                    break
                # MS-OVBA 2.4.1.3.19.1: max(4, ceil(log2(position))).
                offset_bits = max(4, (position - 1).bit_length())
                length_bits = 16 - offset_bits
                maximum = min(len(block) - position, (1 << length_bits) + 2)
                best_length, best_offset = 0, 0
                if maximum >= 3:
                    prefix = block[position:position + 3]
                    candidate = block.rfind(prefix, 0, position + 2)
                    while candidate >= 0:
                        length = 3
                        while length < maximum and block[candidate + length] == block[position + length]:
                            length += 1
                        if length > best_length:
                            best_length, best_offset = length, position - candidate
                        if length == maximum:
                            break
                        candidate = block.rfind(prefix, 0, candidate + 2)
                if best_length >= 3:
                    body[flag_index] |= 1 << bit
                    token = ((best_offset - 1) << length_bits) | (best_length - 3)
                    body += struct.pack('<H', token)
                    position += best_length
                else:
                    body.append(block[position])
                    position += 1
        if len(body) > 4096:
            out += struct.pack('<H', 0x3FFF) + block.ljust(4096, b' ')
        else:
            out += struct.pack('<H', 0xB000 | (len(body) - 1)) + body
    return bytes(out)


def record(kind, data):
    return struct.pack('<HI', kind, len(data)) + data


def directory_stream(modules, classes=()):
    u32 = lambda value: struct.pack('<I', value)
    u16 = lambda value: struct.pack('<H', value)
    data = record(1, u32(2))  # PROJECTSYSKIND: Macintosh, not Win32/Win64.
    data += record(2, u32(0x409)) + record(0x14, u32(0x409))
    data += record(3, u16(1252)) + record(4, b'RadiusNative')
    data += record(5, b'') + record(0x40, b'')
    data += record(6, b'') + record(0x3D, b'')
    data += record(7, u32(0)) + record(8, u32(0))
    data += struct.pack('<HIIH', 9, 4, 1, 0)
    data += record(0x0C, b'') + record(0x3C, b'')
    references = [
        ('stdole', '{00020430-0000-0000-C000-000000000046}', '2.0', 'OLE Automation'),
        ('PowerPoint', '{91493440-5A91-11CF-8700-00AA0060263B}', '2.12', 'Microsoft PowerPoint 16.0 Object Library'),
        ('Office', '{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', '2.8', 'Microsoft Office 16.0 Object Library'),
    ]
    for name, guid, version, description in references:
        # Resolve by installed library GUID; no Windows drive/path assumptions.
        libid = ('*\\G' + guid + '#' + version + '#0##' + description).encode('cp1252')
        data += record(0x16, name.encode('cp1252')) + record(0x3E, name.encode('utf-16le'))
        data += record(0x0D, u32(len(libid)) + libid + u32(0) + u16(0))
    data += record(0x0F, u16(len(modules))) + record(0x13, u16(0))
    for name in modules:
        encoded = name.encode('cp1252')
        unicode = name.encode('utf-16le')
        data += record(0x19, encoded) + record(0x47, unicode)
        data += record(0x1A, encoded) + record(0x32, unicode)
        data += record(0x1C, b'') + record(0x48, b'')
        data += record(0x31, u32(0)) + record(0x1E, u32(0)) + record(0x2C, u16(0))
        data += struct.pack('<HI', 0x22 if name in classes else 0x21, 0)
        if name in classes:
            data += struct.pack('<HI', 0x28, 0)  # MODULEPRIVATE for a private class.
        data += struct.pack('<HI', 0x2B, 0)
    data += struct.pack('<HI', 0x10, 0)
    return compressed(data)


def compound_file(streams):
    """Small CFB v3 writer with FAT, MiniFAT and balanced directory trees."""
    entries = [dict(name='Root Entry', kind=5), dict(name='VBA', kind=1)]
    for path, data in streams.items():
        entries.append(dict(name=path.split('/')[-1], path=path, data=data, kind=2))
    for entry in entries:
        entry.update(left=FREE, right=FREE, child=FREE, color=1, start=END, size=0)

    def sibling_tree(ids):
        ordered = sorted(ids, key=lambda i: (len(entries[i]['name']), entries[i]['name'].upper()))
        maximum = math.floor(math.log2(len(ids))) if ids else 0

        def build(items, depth=0):
            if not items:
                return FREE
            middle = len(items) // 2
            index = items[middle]
            entries[index]['left'] = build(items[:middle], depth + 1)
            entries[index]['right'] = build(items[middle + 1:], depth + 1)
            entries[index]['color'] = 0 if depth == maximum and depth > 0 else 1
            return index
        return build(ordered)

    entries[0]['child'] = sibling_tree([1] + [i for i, e in enumerate(entries) if e.get('path') and '/' not in e['path']])
    entries[1]['child'] = sibling_tree([i for i, e in enumerate(entries) if e.get('path', '').startswith('VBA/')])
    sectors, fat, mini, mini_fat = [], [], bytearray(), []

    def allocate(data):
        start = len(sectors)
        count = (len(data) + 511) // 512
        for i in range(count):
            sectors.append(data[i * 512:(i + 1) * 512].ljust(512, b'\x00'))
            fat.append(start + i + 1 if i + 1 < count else END)
        return start if count else END

    for entry in entries[2:]:
        data = entry['data']
        entry['size'] = len(data)
        if len(data) >= 4096:
            entry['start'] = allocate(data)
        else:
            start = len(mini_fat)
            count = (len(data) + 63) // 64
            entry['start'] = start if count else END
            mini.extend(data.ljust(count * 64, b'\x00'))
            mini_fat.extend(start + i + 1 if i + 1 < count else END for i in range(count))
    entries[0]['start'] = allocate(bytes(mini))
    entries[0]['size'] = len(mini)
    first_mini_fat = allocate(b''.join(struct.pack('<I', n) for n in mini_fat).ljust(((len(mini_fat) + 127) // 128) * 512, b'\xff'))
    mini_fat_sectors = (len(mini_fat) + 127) // 128
    directory = bytearray()
    for entry in entries:
        name = (entry['name'] + '\x00').encode('utf-16le')
        if len(name) > 64:
            raise ValueError('CFB directory name too long')
        directory += name.ljust(64, b'\x00')
        directory += struct.pack('<HBBIII', len(name), entry['kind'], entry['color'], entry['left'], entry['right'], entry['child'])
        directory += b'\x00' * 36  # CLSID, state bits, timestamps.
        directory += struct.pack('<IQ', entry['start'], entry['size'])
    first_directory = allocate(bytes(directory))
    fat_count = 1
    while fat_count != (len(sectors) + fat_count + 127) // 128:
        fat_count = (len(sectors) + fat_count + 127) // 128
    if fat_count > 109:
        raise ValueError('Prototype writer does not support external DIFAT sectors')
    fat_ids = list(range(len(sectors), len(sectors) + fat_count))
    fat.extend([FAT] * fat_count)
    fat_data = b''.join(struct.pack('<I', n) for n in fat).ljust(fat_count * 512, b'\xff')
    sectors.extend(fat_data[i * 512:(i + 1) * 512] for i in range(fat_count))
    header = bytearray(512)
    header[:8] = bytes.fromhex('d0cf11e0a1b11ae1')
    struct.pack_into('<HHHHH', header, 24, 0x003E, 3, 0xFFFE, 9, 6)
    struct.pack_into('<IIIIIIIII', header, 40, 0, fat_count, first_directory, 0, 4096, first_mini_fat, mini_fat_sectors, END, 0)
    struct.pack_into('<109I', header, 76, *(fat_ids + [FREE] * (109 - len(fat_ids))))
    return bytes(header) + b''.join(sectors)


def vba_project(modules, classes=()):
    # Unprotected, visible project metadata from the MS-OVBA PROJECT example.
    # The fixed example Project ID matches the obfuscated defaults' ProjectKey.
    project = ['ID="{917DED54-440B-4FD1-A5C1-74ACF261E600}"']
    project += [('Class=' if name in classes else 'Module=') + name for name in modules]
    project += ['Name="RadiusNative"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
                'CMG="0705D8E3D8EDDBF1DBF1DBF1DBF1"', 'DPB="0E0CD1ECDFF4E7F5E7F5E7"', 'GC="1517CAF1D6F9D7F9D706"', '',
                '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '', '[Workspace]']
    project += [name + '=0, 0, 0, 0, C' for name in modules]
    wm = b''.join(name.encode('cp1252') + b'\x00' + name.encode('utf-16le') + b'\x00\x00' for name in modules) + b'\x00\x00'
    streams = {'PROJECT': ('\r\n'.join(project) + '\r\n').encode('cp1252'), 'PROJECTwm': wm,
               'VBA/_VBA_PROJECT': struct.pack('<HHBH', 0x61CC, 0xFFFF, 0, 0), 'VBA/dir': directory_stream(modules, classes)}
    for name, source in modules.items():
        streams['VBA/' + name] = compressed(source.replace('\r\n', '\n').replace('\n', '\r\n').encode('cp1252'))
    return compound_file(streams)


def build(output, native=ROOT / 'native'):
    sources = sorted([*native.glob('*.bas'), *native.glob('*.cls')])
    modules = {path.stem: path.read_text(encoding='utf-8') for path in sources}
    if len(modules) != len(sources):
        raise ValueError('Duplicate VBA module name')
    classes = {path.stem for path in sources if path.suffix == '.cls'}
    if not modules:
        raise ValueError('No VBA sources')
    callbacks = set(re.findall(r'Public Sub (\w+)\(', '\n'.join(modules.values()), re.I))
    ribbon = (native / 'customUI.xml').read_bytes()
    root = ET.fromstring(ribbon)
    for element in root.iter():
        for key, value in element.attrib.items():
            if key.startswith('on') or key.startswith('get'):
                if value not in callbacks:
                    raise ValueError('Missing callback: ' + value)
    images = sorted({e.get('image') for e in root.iter() if e.get('image')})
    image_parts = {}
    relationships = ET.Element('Relationships', xmlns='http://schemas.openxmlformats.org/package/2006/relationships')
    for name in images:
        if not re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', name):
            raise ValueError('Invalid image id: ' + name)
        path = native / 'icons' / (name + '.png')
        if not path.is_file():
            raise ValueError('Missing ribbon image: ' + name)
        payload = path.read_bytes()
        if not payload.startswith(b'\x89PNG\r\n\x1a\n'):
            raise ValueError('Invalid PNG: ' + name)
        image_parts['customUI/images/' + name + '.png'] = payload
        ET.SubElement(relationships, 'Relationship', Id=name,
                      Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                      Target='images/' + name + '.png')
    parts = {
        '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"/></Types>',
        '_rels/.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/></Relationships>',
        'ppt/presentation.xml': '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
        'ppt/_rels/presentation.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
        'ppt/vbaProject.bin': vba_project(modules, classes), 'customUI/customUI.xml': ribbon,
    }
    if images:
        types = ET.fromstring(parts['[Content_Types].xml'])
        ET.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Default',
                      Extension='png', ContentType='image/png')
        parts['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
        parts['customUI/_rels/customUI.xml.rels'] = ET.tostring(relationships, encoding='utf-8', xml_declaration=True)
        parts.update(image_parts)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, (2026, 10, 7, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data)
    return output


def distribution(output, native=ROOT / 'native'):
    """Package a one-time install helper; no executable app or web resources."""
    folder = output.parent / 'RadiusInPptNative-mac'
    folder.mkdir(parents=True, exist_ok=True)
    files = {
        'RadiusInPptNative.ppam': (output, 0o644),
        'Install-RadiusInPptNative.command': (native / 'Install-RadiusInPptNative.command', 0o755),
        'INSTALL.txt': (native / 'INSTALL.txt', 0o644),
    }
    archive_path = output.parent / 'RadiusInPptNative-mac.zip'
    with zipfile.ZipFile(archive_path, 'w') as archive:
        for name, (source, mode) in files.items():
            target = folder / name
            shutil.copyfile(source, target)
            target.chmod(mode)
            info = zipfile.ZipInfo(folder.name + '/' + name, (2026, 10, 7, 0, 0, 0))
            info.create_system = 3
            info.external_attr = (0o100000 | mode) << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, target.read_bytes())
    return archive_path


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, default=ROOT / 'dist' / 'RadiusInPptNative.ppam')
    parser.add_argument('--distribution', action='store_true', help='Also build the Mac installation ZIP')
    args = parser.parse_args()
    output = build(args.output)
    print(output)
    if args.distribution:
        print(distribution(output))
    print('Source-only Mac PPAM built. This build does not run PowerPoint; host validation is separate.')
