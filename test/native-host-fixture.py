"""Build ordinary PPTX fixtures for native Ribbon regression; no test macros.

Requires python-pptx. Run with 'build /tmp/RadiusNativeRegression.pptx', then
exercise the installed Ribbon and save. 'inspect FILE' reads exact saved
geometry, adjustments and tags independently from PowerPoint/VBA.
'verify FILE' checks the final seven-slide protocol in test/README.md.
"""
import json
from pathlib import Path
import posixpath
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
EMU_CM = 360000


def build(path):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.util import Cm

    presentation = Presentation()
    presentation.slide_width, presentation.slide_height = Cm(30), Cm(18)
    tag_number = 0

    def tags(slide, shape, values):
        nonlocal tag_number
        tag_number += 1
        root = ET.Element('{%s}tagLst' % NS['p'])
        for key, value in values.items():
            ET.SubElement(root, '{%s}tag' % NS['p'], name=key.upper(), val=str(value))
        part = Part(PackURI('/ppt/tags/tag%d.xml' % tag_number),
                    'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
                    slide.part.package, ET.tostring(root, encoding='utf-8'))
        relationship = slide.part.relate_to(part, NS['r'] + '/tags')
        nv = shape.element.xpath('./p:nvSpPr/p:nvPr | ./p:nvGrpSpPr/p:nvPr')[0]
        data = OxmlElement('p:custDataLst')
        tag = OxmlElement('p:tags')
        tag.set('{%s}id' % NS['r'], relationship)
        data.append(tag)
        nv.append(data)

    def rounded(container, name, x, y, w, h):
        shape = container.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x), Cm(y), Cm(w), Cm(h))
        shape.name = name
        shape.adjustments[0] = 0.15
        return shape

    def slide():
        return presentation.slides.add_slide(presentation.slide_layouts[6])

    # Slide 1: layout parent + child + plain shape + arrow, all selected.
    s = slide()
    parent = rounded(s.shapes, 'Parent', 2, 2, 9, 5)
    child = rounded(s.shapes, 'Child', 13, 2, 4, 3)
    rounded(s.shapes, 'Plain', 19, 2, 5, 4)
    s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(2), Cm(10), Cm(5), Cm(2)).name = 'Arrow'
    tags(s, parent, {'layoutParent_v1': json.dumps({'rows': 1, 'cols': 1, 'padding': 0.35, 'linkRMode': 'same', 'childIds': [str(child.shape_id)]}), 'radiusLock_v1': '0.75'})
    tags(s, child, {'layoutChild_v1': str(parent.shape_id), 'radiusLock_v1': '0.45', 'custom': 'KeepCase'})

    # Slide 2: the final shape is strict; no earlier shape may be mutated.
    s = slide()
    rounded(s.shapes, 'EarlyPlain', 2, 2, 5, 3)
    child = rounded(s.shapes, 'LateProtectedChild', 10, 2, 5, 3)
    tags(s, child, {'layoutChild_v1': '42', 'radiusLockStrict_v1': '1', 'radiusLock_v1': '0.45'})

    # Slide 3: protection buttons on a tagged leaf.
    s = slide()
    child = rounded(s.shapes, 'ProtectionChild', 2, 2, 6, 4)
    tags(s, child, {'layoutChild_v1': '42', 'radiusLock_v1': '0.60', 'custom': 'KeepCase'})

    # Slide 4: 4x2cm; 20%=0.4cm, zero=0, upper clamp=1cm.
    s = slide()
    rounded(s.shapes, 'UnitsChild', 2, 2, 4, 2)

    # Slides 5/6: nested scaled groups, with and without strict leaf.
    for protected in (False, True):
        s = slide()
        outer = s.shapes.add_group_shape()
        outer.name = 'OuterGroup'
        parent = rounded(outer.shapes, 'GroupParent', 2, 2, 6, 4)
        inner = outer.shapes.add_group_shape()
        inner.name = 'InnerGroup'
        child = rounded(inner.shapes, 'NestedChild', 10, 2, 4, 3)
        arrow = inner.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(10), Cm(7), Cm(4), Cm(2))
        arrow.name = 'NestedArrow'
        # Refresh enclosing bounds after populating the nested shape tree.
        outer.element.recalculate_extents()
        tags(s, parent, {'layoutParent_v1': json.dumps({'rows': 1, 'cols': 1, 'linkRMode': 'same', 'childIds': [str(child.shape_id)]})})
        values = {'layoutChild_v1': str(parent.shape_id), 'custom': 'KeepCase'}
        if protected:
            values['radiusLockStrict_v1'] = '1'
        tags(s, child, values)
        tags(s, inner, {'customInner': 'KeepInnerCase'})
        tags(s, outer, {'customOuter': 'KeepOuterCase'})
        outer.width = int(outer.width * 1.4)
        outer.height = int(outer.height * 0.8)

    # Slide 7: no eligible rounded rectangle.
    s = slide()
    s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(2), Cm(2), Cm(5), Cm(2)).name = 'OnlyArrow'
    presentation.save(path)


def inspect(path, slide_count=7):
    result = {}
    with zipfile.ZipFile(path) as archive:
        for i in range(1, slide_count + 1):
            slide_path = 'ppt/slides/slide%d.xml' % i
            root = ET.fromstring(archive.read(slide_path))
            rel_path = 'ppt/slides/_rels/slide%d.xml.rels' % i
            relationships = {r.get('Id'): posixpath.normpath('ppt/slides/' + r.get('Target'))
                             for r in ET.fromstring(archive.read(rel_path))}
            items = {}

            def visit(tree, transform=(1, 1, 0, 0), parent_name=None):
                sx, sy, dx, dy = transform
                for shape in tree:
                    group = shape.tag == '{%s}grpSp' % NS['p']
                    if not group and shape.tag != '{%s}sp' % NS['p']:
                        continue
                    nv = shape.find('p:nvGrpSpPr' if group else 'p:nvSpPr', NS)
                    identity = nv.find('p:cNvPr', NS)
                    xfrm = shape.find('p:grpSpPr/a:xfrm' if group else 'p:spPr/a:xfrm', NS)
                    off, ext = xfrm.find('a:off', NS), xfrm.find('a:ext', NS)
                    x, y, w, h = (int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy')))
                    box = [(x * sx + dx) / EMU_CM, (y * sy + dy) / EMU_CM, w * sx / EMU_CM, h * sy / EMU_CM]
                    entry = {'id': identity.get('id'), 'group': group, 'parent': parent_name, 'boxCm': box, 'tags': {}}
                    tag = nv.find('p:nvPr/p:custDataLst/p:tags', NS)
                    if tag is not None:
                        part = relationships[tag.get('{%s}id' % NS['r'])]
                        entry['tags'] = {t.get('name').upper(): t.get('val') for t in ET.fromstring(archive.read(part))}
                    geom = shape.find('p:spPr/a:prstGeom', NS)
                    if geom is not None and geom.get('prst') == 'roundRect':
                        adjustment = geom.find('a:avLst/a:gd', NS)
                        fraction = float(adjustment.get('fmla').split()[-1]) / 100000 if adjustment is not None else 0.16667
                        entry['fraction'] = fraction
                        entry['radiusCm'] = fraction * min(box[2:])
                    items[identity.get('name')] = entry
                    if group:
                        ch_off, ch_ext = xfrm.find('a:chOff', NS), xfrm.find('a:chExt', NS)
                        gsx, gsy = w / int(ch_ext.get('cx')), h / int(ch_ext.get('cy'))
                        visit(shape, (sx * gsx, sy * gsy, dx + sx * (x - int(ch_off.get('x')) * gsx), dy + sy * (y - int(ch_off.get('y')) * gsy)), identity.get('name'))
            visit(root.find('p:cSld/p:spTree', NS))
            result[str(i)] = items
    return result


def verify(path, group_path=None):
    """Assert saved host outcomes, not VBA source or a mock implementation."""
    with tempfile.TemporaryDirectory() as directory:
        baseline = Path(directory) / 'baseline.pptx'
        build(baseline)
        before = inspect(baseline)
    after = inspect(path)
    sources = [{'file': str(path), 'slides': list(before)}]
    if group_path is not None:
        # A separate fresh fixture keeps group transactions isolated from
        # manual movement of the first interactive test document.
        group_state = inspect(group_path)
        for slide in before:
            if slide != '5':
                assert group_state[slide] == before[slide], f'Unexpected edit in group fixture: slide {slide}'
        after['5'] = group_state['5']
        sources = [{'file': str(path), 'slides': ['1', '2', '3', '4', '6', '7']},
                   {'file': str(group_path), 'slides': ['5']}]
    targets = {
        ('1', 'Parent'): 0.5, ('1', 'Child'): 0.5, ('1', 'Plain'): 0.5,
        ('3', 'ProtectionChild'): 0.3, ('4', 'UnitsChild'): 1.0,
        ('5', 'GroupParent'): 0.3, ('5', 'NestedChild'): 0.3,
    }
    locks = {('1', 'Parent'), ('1', 'Child'), ('3', 'ProtectionChild'),
             ('5', 'GroupParent'), ('5', 'NestedChild')}
    for slide, shapes in before.items():
        assert shapes.keys() == after[slide].keys(), f'Shape names/count changed on slide {slide}'
        for name, original in shapes.items():
            actual = after[slide][name]
            context = f'{slide}/{name}'
            assert actual['group'] == original['group'], f'Group type changed: {context}'
            assert actual['parent'] == original['parent'], f'Group hierarchy changed: {context}'
            assert all(abs(x - y) < 0.0001 for x, y in zip(actual['boxCm'], original['boxCm'])), f'Geometry changed: {context}'
            if not original['group'] or slide in ('2', '6', '7'):
                assert actual['id'] == original['id'], f'Leaf/protected identity changed: {context}'
            expected_tags = dict(original['tags'])
            key = (slide, name)
            if key in locks:
                expected_tags['RADIUSLOCK_V1'] = f'{targets[key]:.6f}'
            assert actual['tags'] == expected_tags, f'Tags changed unexpectedly: {context}'
            if 'radiusCm' in original:
                target = targets.get(key, original['radiusCm'])
                assert abs(actual['radiusCm'] - target) < 0.0001, f'Radius mismatch: {context}'
    checks = [
        'tagged parent/child + mixed multiselection: 0.50cm',
        'late strict leaf: entire selection unchanged',
        'tagged leaf: protect/reject/unprotect/apply 0.30cm',
        '4x2cm shape: upper radius clamp 1.00cm',
        'scaled nested group: 0.30cm + protect/reject/unprotect; hierarchy/tags/geometry preserved',
        'strict nested leaf: group and all descendants unchanged',
        'non-rounded arrow: unchanged',
    ]
    if group_path is not None:
        checks = [
            'three rounded shapes: 0.30 + two up steps = 0.50cm; arrow preserved',
            'mixed strict selection: partial-protection state, radius controls disabled, unchanged',
            'zero + three up steps = 0.30cm; protection toggle on/off, metadata preserved',
            '20% up/down intermediate values and 50% upper limit; final radius 1.00cm',
            'fresh scaled nested group: 0.10 + two up steps = 0.30cm; protection on/off, geometry and hierarchy preserved',
            'strict nested group: partial-protection state, writes disabled, unchanged',
            'non-rounded arrow: no eligible radius, writes disabled, unchanged',
        ]
    return {'passed': len(checks), 'checks': checks, 'sourceFiles': sources, 'savedState': after}


if __name__ == '__main__':
    command, path = sys.argv[1:3]
    if command == 'build':
        build(path)
        print(path)
    elif command == 'inspect':
        print(json.dumps(inspect(path), ensure_ascii=False, indent=2))
    elif command == 'verify':
        print(json.dumps(verify(path), ensure_ascii=False, indent=2))
    elif command == 'verify-controls':
        print(json.dumps(verify(path, sys.argv[3]), ensure_ascii=False, indent=2))
    else:
        raise ValueError('Use build, inspect, verify or verify-controls (with a separate group fixture).')
