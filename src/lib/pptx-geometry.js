/* Read preset geometry from a single-slide PPTX. No Office.js calls here.
 * Shape.type only distinguishes GeometricShape; adjustment counts do not
 * distinguish roundRect from arrows, crosses, etc. Unknown geometry is skipped.
 */
(function (root) {
  function parseSlideGeometry(xml) {
    const result = new Map();
    const stack = [];
    let shape = null;
    const attr = (text, name) => {
      const match = text.match(new RegExp('(?:^|\\s)' + name + '\\s*=\\s*["\x27]([^"\x27]*)["\x27]'));
      return match ? match[1] : null;
    };
    // OOXML is well-formed XML. Only read sp/cNvPr/prstGeom attributes; no
    // markup is rendered or executed. Namespace prefixes are not assumed.
    const tokens = String(xml).match(/<!--[^]*?-->|<[^>]+>/g) || [];
    for (const token of tokens) {
      if (/^<\?|^<!/.test(token)) continue;
      const match = token.match(/^<(\/?)\s*(?:[\w.-]+:)?([\w.-]+)/);
      if (!match) continue;
      const closing = !!match[1];
      const name = match[2];
      if (closing) {
        if (name === 'sp' && shape && stack.length === shape.depth) {
          if (shape.id != null) result.set(shape.id, shape.preset);
          shape = null;
        }
        stack.pop();
        continue;
      }
      const parent = stack[stack.length - 1];
      stack.push(name);
      if (name === 'sp') shape = { depth: stack.length, id: null, preset: null };
      if (shape && name === 'cNvPr' && parent === 'nvSpPr') shape.id = attr(token, 'id');
      if (shape && name === 'prstGeom' && parent === 'spPr') shape.preset = attr(token, 'prst');
      if (/\/\s*>$/.test(token)) stack.pop();
    }
    return result;
  }

  function readSlideGeometry(base64) {
    const zip = typeof module !== 'undefined' && module.exports
      ? require('./vendor/fflate-0.8.2.js') : root.fflate;
    if (!zip) throw new Error('PPTX decoder unavailable');
    const binary = atob(base64);
    const bytes = Uint8Array.from(binary, (ch) => ch.charCodeAt(0));
    const entries = zip.unzipSync(bytes, {
      filter: (entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.name) && entry.originalSize <= 8 * 1024 * 1024,
    });
    const names = Object.keys(entries);
    if (names.length !== 1) throw new Error('Expected one exported slide');
    return parseSlideGeometry(zip.strFromU8(entries[names[0]]));
  }
  const api = { parseSlideGeometry, readSlideGeometry };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.PptxGeometry = api;
})(typeof window !== 'undefined' ? window : globalThis);
