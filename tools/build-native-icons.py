#!/usr/bin/env python3
"""Render the project's own vector icons as embedded Ribbon PNGs.

Pillow is needed only to regenerate these assets, never to build/use the PPAM.
Each primitive is also saved to SVG for future editing.
"""
from pathlib import Path
from html import escape
from PIL import Image, ImageDraw

DESTINATION = Path(__file__).resolve().parent.parent / 'native' / 'icons'
BLUE, LIGHT, GREEN, PURPLE = '#2261ac', '#edf4fd', '#248052', '#7554a3'
SCALE = 4


class Icon:
    def __init__(self):
        self.image = Image.new('RGBA', (32 * SCALE, 32 * SCALE))
        self.draw = ImageDraw.Draw(self.image)
        self.elements = []

    def line(self, points, color=BLUE, width=2.2):
        xy = [(round(x * SCALE), round(y * SCALE)) for x, y in points]
        self.draw.line(xy, fill=color, width=round(width * SCALE), joint='curve')
        radius = width * SCALE / 2
        for x, y in (xy[0], xy[-1]):
            self.draw.ellipse((x-radius, y-radius, x+radius, y+radius), fill=color)
        self.elements.append('<polyline points="%s" fill="none" stroke="%s" stroke-width="%s" stroke-linecap="round" stroke-linejoin="round"/>' % (' '.join(f'{x},{y}' for x, y in points), color, width))

    def rect(self, box, radius=0, color=BLUE, fill=LIGHT, width=2.2):
        self.draw.rounded_rectangle(tuple(round(p * SCALE) for p in box), radius=round(radius * SCALE), outline=color, fill=fill, width=round(width * SCALE))
        x, y, right, bottom = box
        self.elements.append(f'<rect x="{x}" y="{y}" width="{right-x}" height="{bottom-y}" rx="{radius}" stroke="{color}" fill="{fill}" stroke-width="{width}"/>')

    def polygon(self, points, fill=BLUE):
        self.draw.polygon([(round(x * SCALE), round(y * SCALE)) for x, y in points], fill=fill)
        self.elements.append('<polygon points="%s" fill="%s"/>' % (' '.join(f'{x},{y}' for x, y in points), fill))

    def circle(self, center, radius, color=BLUE, fill=LIGHT, width=2.2):
        x, y = center
        self.draw.ellipse(tuple(round(v * SCALE) for v in (x-radius,y-radius,x+radius,y+radius)), outline=color, fill=fill, width=round(width*SCALE))
        self.elements.append(f'<circle cx="{x}" cy="{y}" r="{radius}" stroke="{color}" fill="{fill}" stroke-width="{width}"/>')

    def save(self, name):
        self.image.resize((32, 32), Image.Resampling.LANCZOS).save(DESTINATION / (name + '.png'))
        (DESTINATION / (name + '.svg')).write_text('<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32"><title>' + escape(name) + '</title>' + ''.join(self.elements) + '</svg>\n')


def main():
    DESTINATION.mkdir(parents=True, exist_ok=True)
    for name, radius in (('presetZero',0), ('presetSmall',3), ('presetMedium',6), ('presetLarge',9)):
        icon = Icon()
        icon.rect((5,5,27,27), radius)
        icon.save(name)
    icon = Icon()
    icon.rect((4,4,27,27), 8)
    icon.circle((24,24),7,color=GREEN,fill='#ffffff',width=0)
    icon.line([(19,24),(23,28),(30,19)],GREEN,2.8)
    icon.save('applyRadius')
    for name, points in (('stepUp',[(8,21),(16,12),(24,21)]),('stepDown',[(8,12),(16,21),(24,12)])):
        icon = Icon()
        icon.line(points,width=3)
        icon.save(name)
    icon = Icon()
    icon.rect((3,4,23,24),6)
    icon.circle((20,20),7,fill='#ffffff')
    icon.line([(25,25),(29,29)],width=3)
    icon.line([(16,20),(24,20)],width=1.8)
    icon.save('readSelection')
    icon = Icon()
    shield = [(16,3),(27,7),(26,18),(22,24),(16,29),(10,24),(6,18),(5,7),(16,3)]
    icon.polygon(shield, LIGHT)
    icon.line(shield,width=2.3)
    icon.line([(11,16),(15,20),(22,12)],GREEN,2.8)
    icon.save('protect')
    icon = Icon()
    icon.line([(22,13),(22,9),(20,5),(15,4),(10,7),(10,16)],width=2.6)
    icon.rect((6,14,26,28),3)
    icon.circle((16,20),1.8,color=BLUE,fill=BLUE,width=0)
    icon.line([(16,21),(16,24)],width=2)
    icon.save('unprotect')
    icon = Icon()
    icon.polygon([(12,4),(20,4),(20,13),(28,26),(26,29),(6,29),(4,26),(12,13)],'#f2eef8')
    icon.line([(10,4),(22,4)],PURPLE)
    icon.line([(12,5),(12,13),(5,25),(6,28),(26,28),(27,25),(20,13),(20,5)],PURPLE)
    icon.line([(10,21),(22,21)],PURPLE)
    icon.circle((15,24),1.3,color=PURPLE,fill=PURPLE,width=0)
    icon.save('selfTest')


if __name__ == '__main__':
    main()
