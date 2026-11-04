/* Editable DOCX export. Runs locally in the browser without external dependencies. */
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.MCWordExport = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const encoder = new TextEncoder();
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const xmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  function cleanText(text) {
    return String(text).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');
  }
  function escapeXml(text) {
    return text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
  }
  function wordFilename(text) {
    const first = cleanText(text).split('\n').find(line => line.trim());
    if (!first) return '周会司仪稿.docx';
    const safe = Array.from(first.trim().replace(/[<>:"/\\|?*]/g, '').replace(/[. ]+$/g, '')).slice(0, 80).join('');
    return safe ? safe + '_司仪稿.docx' : '周会司仪稿.docx';
  }
  function isIntroduction(line) {
    return /[：:]$/.test(line) || /代表同学有[：:]?$/.test(line) || /^我校.*代表$/.test(line) ||
      /^(在.+中[，,：:]|获奖班级有[：:]|现在进行颁奖仪式|接下来[，,].*(上台颁奖|为我们颁奖)|接下来继续由)/.test(line);
  }
  function documentXml(text) {
    const lines = text.split('\n');
    const title = lines.findIndex(line => line.trim());
    const keep = lines.map((line, i) => i === title || isIntroduction(line.trim()));
    for (let i = 0; i < lines.length; i++) {
      if (!lines[i].trim()) continue;
      let next = i + 1;
      while (next < lines.length && !lines[next].trim()) next++;
      if (next < lines.length && /一同合影留念/.test(lines[next])) keep[i] = true;
    }
    let previous = -1;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].trim()) previous = i;
      else if (previous >= 0) keep[i] = keep[previous];
    }
    const paragraphs = lines.map((line, i) => {
      const style = i === title ? 'Title' : line.trim() ? 'Normal' : 'Blank';
      const runs = line.split('\t').map((part, j) =>
        (j ? '<w:r><w:tab/></w:r>' : '') + '<w:r><w:t xml:space="preserve">' + escapeXml(part) + '</w:t></w:r>').join('');
      return '<w:p><w:pPr><w:pStyle w:val="' + style + '"/>' +
        (keep[i] ? '<w:keepNext/>' : '') + '<w:keepLines/></w:pPr>' + runs + '</w:p>';
    }).join('');
    return xmlHeader + '<w:document xmlns:w="' + ns + '"><w:body>' + paragraphs +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>';
  }
  function stylesXml() {
    const run = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="KaiTi" w:cs="Times New Roman"/>' +
      '<w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>';
    return xmlHeader + '<w:styles xmlns:w="' + ns + '">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' + run + '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl/><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr>' + run + '</w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:color w:val="000000"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Blank"><w:name w:val="Blank"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="180" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr></w:style></w:styles>';
  }

  // STORE-mode ZIP entries keep the document portable with no compression library.
  const crcTable = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let bit = 0; bit < 8; bit++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[i] = c >>> 0;
  }
  function crc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }
  function zip(entries) {
    const local = [], central = [];
    let offset = 0, centralSize = 0;
    for (const [path, content] of entries) {
      const name = encoder.encode(path), data = encoder.encode(content), crc = crc32(data);
      const head = new Uint8Array(30 + name.length), v = new DataView(head.buffer);
      v.setUint32(0, 0x04034b50, true);v.setUint16(4, 20, true);v.setUint16(6, 0x0800, true);
      v.setUint16(12, 33, true); // 1980-01-01, a valid deterministic DOS date.
      v.setUint32(14, crc, true);v.setUint32(18, data.length, true);v.setUint32(22, data.length, true);v.setUint16(26, name.length, true);
      head.set(name, 30);local.push(head, data);
      const dir = new Uint8Array(46 + name.length), d = new DataView(dir.buffer);
      d.setUint32(0, 0x02014b50, true);d.setUint16(4, 20, true);d.setUint16(6, 20, true);d.setUint16(8, 0x0800, true);d.setUint16(14, 33, true);
      d.setUint32(16, crc, true);d.setUint32(20, data.length, true);d.setUint32(24, data.length, true);d.setUint16(28, name.length, true);d.setUint32(42, offset, true);
      dir.set(name, 46);central.push(dir);centralSize += dir.length;offset += head.length + data.length;
    }
    const end = new Uint8Array(22), e = new DataView(end.buffer);
    e.setUint32(0, 0x06054b50, true);e.setUint16(8, entries.length, true);e.setUint16(10, entries.length, true);e.setUint32(12, centralSize, true);e.setUint32(16, offset, true);
    const result = new Uint8Array(offset + centralSize + end.length);
    let cursor = 0;
    for (const chunk of [...local, ...central, end]) { result.set(chunk, cursor);cursor += chunk.length; }
    return result;
  }
  function buildDocx(source) {
    const text = cleanText(source);
    if (!text.trim()) throw new Error('请先生成或输入司仪稿。');
    const title = text.split('\n').find(line => line.trim()).trim();
    const created = new Date().toISOString();
    return zip([
      ['[Content_Types].xml', xmlHeader + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'],
      ['_rels/.rels', xmlHeader + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + rel + 'officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="' + rel + 'extended-properties" Target="docProps/app.xml"/></Relationships>'],
      ['word/document.xml', documentXml(text)],
      ['word/styles.xml', stylesXml()],
      ['word/_rels/document.xml.rels', xmlHeader + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + rel + 'styles" Target="styles.xml"/></Relationships>'],
      ['docProps/core.xml', xmlHeader + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + escapeXml(title) + '</dc:title><dc:creator>华仁周会司仪稿生成器</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + created + '</dcterms:created></cp:coreProperties>'],
      ['docProps/app.xml', xmlHeader + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>华仁周会司仪稿生成器</Application></Properties>'],
    ]);
  }
  return Object.freeze({ buildDocx, wordFilename });
});
