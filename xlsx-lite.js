/* Minimal .xlsx reader for the browser - no libraries.
   readXlsx(arrayBuffer) -> { "Sheet name": [[cell, cell, ...], ...], ... }
   Cells are numbers, strings, booleans or null. Dates stay as Excel serial numbers. */
(function () {
  'use strict';

  function u16(v, o) { return v.getUint16(o, true); }
  function u32(v, o) { return v.getUint32(o, true); }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('This browser is too old to read Excel files. Update Chrome or Safari.');
    }
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function unzipIndex(buf) {
    const v = new DataView(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
      if (u32(v, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('This file is not an .xlsx workbook.');
    const count = u16(v, eocd + 10);
    let p = u32(v, eocd + 16);
    const files = {};
    const dec = new TextDecoder();
    for (let n = 0; n < count; n++) {
      if (u32(v, p) !== 0x02014b50) throw new Error('The workbook file looks damaged.');
      const method = u16(v, p + 10);
      const csize = u32(v, p + 20);
      const nameLen = u16(v, p + 28), extraLen = u16(v, p + 30), commentLen = u16(v, p + 32);
      const local = u32(v, p + 42);
      const name = dec.decode(new Uint8Array(buf, p + 46, nameLen));
      files[name] = { method, csize, local };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return files;
  }

  async function readEntry(buf, files, name) {
    const f = files[name] || files[name.replace(/^\//, '')];
    if (!f) return null;
    const v = new DataView(buf);
    const start = f.local + 30 + u16(v, f.local + 26) + u16(v, f.local + 28);
    const raw = new Uint8Array(buf, start, f.csize);
    const bytes = f.method === 0 ? raw : await inflateRaw(raw);
    return new TextDecoder().decode(bytes);
  }

  const xml = (s) => new DOMParser().parseFromString(s, 'application/xml');
  const all = (node, tag) => Array.from(node.getElementsByTagNameNS('*', tag));
  const first = (node, tag) => node.getElementsByTagNameNS('*', tag)[0] || null;

  function colIndex(ref) {
    let n = 0;
    for (const ch of ref) {
      const c = ch.charCodeAt(0);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  async function readXlsx(buf) {
    const files = unzipIndex(buf);
    const wb = xml(await readEntry(buf, files, 'xl/workbook.xml'));
    const rels = xml(await readEntry(buf, files, 'xl/_rels/workbook.xml.rels'));
    const relTarget = {};
    all(rels, 'Relationship').forEach((r) => { relTarget[r.getAttribute('Id')] = r.getAttribute('Target'); });

    const shared = [];
    const ssText = await readEntry(buf, files, 'xl/sharedStrings.xml');
    if (ssText) {
      all(xml(ssText), 'si').forEach((si) => {
        shared.push(all(si, 't').filter((t) => !t.closest || !t.closest('rPh')).map((t) => t.textContent).join(''));
      });
    }

    const out = {};
    for (const s of all(wb, 'sheet')) {
      const name = s.getAttribute('name');
      const rid = s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id');
      let target = relTarget[rid];
      if (!target) continue;
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const text = await readEntry(buf, files, target);
      if (!text) continue;
      const doc = xml(text);
      const rows = [];
      all(doc, 'row').forEach((row) => {
        const r = parseInt(row.getAttribute('r'), 10) - 1;
        const cells = [];
        let auto = 0;
        all(row, 'c').forEach((c) => {
          const ref = c.getAttribute('r');
          const ci = ref ? colIndex(ref) : auto;
          auto = ci + 1;
          const t = c.getAttribute('t');
          const vEl = first(c, 'v');
          let val = null;
          if (t === 's') val = vEl ? shared[parseInt(vEl.textContent, 10)] ?? '' : '';
          else if (t === 'inlineStr') { const is = first(c, 'is'); val = is ? all(is, 't').map((x) => x.textContent).join('') : ''; }
          else if (t === 'str') val = vEl ? vEl.textContent : '';
          else if (t === 'b') val = vEl ? vEl.textContent === '1' : null;
          else if (t === 'e') val = null;
          else if (vEl && vEl.textContent !== '') val = Number(vEl.textContent);
          cells[ci] = val;
        });
        rows[isNaN(r) ? rows.length : r] = cells;
      });
      for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
      out[name] = rows;
    }
    return out;
  }

  window.readXlsx = readXlsx;
})();
