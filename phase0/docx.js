/*
 * สร้างไฟล์ .docx หัวรายงานการประชุม 1 หน้า (ระยะ 0)
 * เขียน WordprocessingML เอง แล้วประกอบด้วย MiniZip (zip.js)
 *
 * หัวใจของภาษาไทยใน Word: ตัวอักษรไทยเป็น "complex script"
 * Word จึงใช้ค่า w:rFonts/@w:cs, w:szCs, w:bCs และ w:lang/@w:bidi
 * ถ้าใส่แค่ฟอนต์ ascii จะได้ฟอนต์ไทยผิดตัวและตัดคำผิด
 */
(function (global) {
  'use strict';

  // ชื่อ family จริงในไฟล์ฟอนต์ TH SarabunNew (name table ID 1) คือ "TH Sarabun New"
  var FONT = 'TH Sarabun New';
  var CM = 567; // twips ต่อ 1 ซม.
  var NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  // ค่าหน้ากระดาษ ประมาณจาก PDF ตัวอย่าง (รอเทียบกับไฟล์ Word ต้นฉบับ)
  var PAGE = {
    width: 11906, height: 16838,           // A4
    top: 2.0, bottom: 2.0, left: 3.0, right: 2.0, footer: 1.0, header: 1.25, // ซม.
    fontSize: 16, footerSize: 14,          // pt
    indent: 0.75,                          // ซม. ระยะเยื้องชื่อในรายชื่อ
    roleTab: 11.2                          // ซม. จากขอบซ้ายของเนื้อหา ถึงคอลัมน์ตำแหน่ง
  };

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function tw(cm) { return Math.round(cm * CM); }

  var THAI_SPLIT = /([฀-๿]+)/;
  var THAI = /^[฀-๿]/;

  // text อาจมี \t (กระโดดไป tab stop) และ \n (ขึ้นบรรทัดใหม่ในย่อหน้าเดิม)
  // ข้อความไทยแยกเป็น run ของตัวเองและติด <w:cs/> (complex script)
  // ถ้าไม่ติด Word จะไม่ใช้พจนานุกรมตัดคำไทย และตัดบรรทัดกลางคำ
  // (ทดสอบแล้ว: ตั้ง w:lang เป็น th-TH อย่างเดียวไม่พอ)
  function run(text, o) {
    o = o || {};
    var base = '';                                      // ลำดับตาม schema: b, sz, u, cs
    if (o.b) base += '<w:b/><w:bCs/>';
    if (o.size) base += '<w:sz w:val="' + o.size * 2 + '"/><w:szCs w:val="' + o.size * 2 + '"/>';
    if (o.u) base += '<w:u w:val="single"/>';
    var out = '';
    function emit(rPr, body) {
      out += '<w:r>' + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') + body + '</w:r>';
    }
    String(text).split('\n').forEach(function (line, i) {
      if (i > 0) emit(base, '<w:br/>');
      line.split('\t').forEach(function (part, j) {
        if (j > 0) emit(base, '<w:tab/>');
        part.split(THAI_SPLIT).forEach(function (seg) {
          if (!seg) return;
          emit(THAI.test(seg) ? base + '<w:cs/>' : base,
            '<w:t xml:space="preserve">' + esc(seg) + '</w:t>');
        });
      });
    });
    return out;
  }

  function para(runs, o) {
    o = o || {};
    var pPr = '';                                       // ลำดับตาม schema: pStyle, keepNext, tabs, spacing, ind, jc
    if (o.style) pPr += '<w:pStyle w:val="' + o.style + '"/>';
    if (o.keepNext) pPr += '<w:keepNext/>';
    if (o.tabs) {
      pPr += '<w:tabs>' + o.tabs.map(function (t) {
        return '<w:tab w:val="' + (t.val || 'left') + '" w:pos="' + tw(t.pos) + '"/>';
      }).join('') + '</w:tabs>';
    }
    if (o.before || o.after) {
      pPr += '<w:spacing w:before="' + (o.before || 0) * 20 + '" w:after="' + (o.after || 0) * 20 + '"/>';
    }
    if (o.ind) {
      pPr += '<w:ind w:left="' + tw(o.ind.left || 0) + '"' +
        (o.ind.hanging ? ' w:hanging="' + tw(o.ind.hanging) + '"' : '') +
        (o.ind.firstLine ? ' w:firstLine="' + tw(o.ind.firstLine) + '"' : '') + '/>';
    }
    if (o.align) pPr += '<w:jc w:val="' + o.align + '"/>';
    var content = typeof runs === 'string' ? (runs ? run(runs) : '') : runs.join('');
    return '<w:p>' + (pPr ? '<w:pPr>' + pPr + '</w:pPr>' : '') + content + '</w:p>';
  }

  function field(instr, cached, size) {
    var rPr = '<w:rPr><w:sz w:val="' + size * 2 + '"/><w:szCs w:val="' + size * 2 + '"/></w:rPr>';
    return '<w:r>' + rPr + '<w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r>' + rPr + '<w:instrText xml:space="preserve"> ' + instr + ' </w:instrText></w:r>' +
      '<w:r>' + rPr + '<w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r>' + rPr + '<w:t>' + cached + '</w:t></w:r>' +
      '<w:r>' + rPr + '<w:fldChar w:fldCharType="end"/></w:r>';
  }

  // ---------- เนื้อหาหน้า 1 ตามแบบรายงานสภาวิชาการ ----------
  function buildBody(d, opt) {
    var out = [];
    var roleTabs = [{ pos: PAGE.roleTab }];
    var listInd = { left: PAGE.indent, hanging: PAGE.indent };

    function heading(text) {
      out.push(para([run(text, { b: true, u: true })], { before: 6, keepNext: true }));
    }
    function numbered(list, withRole) {
      list.forEach(function (p, i) {
        var text = (i + 1) + '.\t' + p.name + (withRole && p.role ? '\t' + p.role : '');
        if (p.note) text += '\n' + p.note;
        out.push(para([run(text)], { ind: listInd, tabs: roleTabs }));
      });
    }

    // หัวรายงาน: 4 บรรทัดตัวหนาจัดกลาง + เส้นคั่น
    d.header.forEach(function (line) {
      out.push(para([run(line, { b: true })], { align: 'center' }));
    });
    out.push(para([run('--------------------------')], { align: 'center' }));

    heading('ผู้มาประชุม');
    (d.advisors || []).forEach(function (p) {
      out.push(para([run(p.name + '\t' + p.role)], { ind: { left: PAGE.indent }, tabs: roleTabs }));
    });
    numbered(d.members || [], true);

    if (d.absent && d.absent.length) { heading('ลาประชุม'); numbered(d.absent, true); }
    if (d.attendees && d.attendees.length) { heading('ผู้เข้าร่วมประชุม'); numbered(d.attendees, false); }

    out.push(para([run('เริ่มประชุมเวลา ' + d.startTime + ' น.', { b: true })], { before: 12 }));
    out.push(para([run('เมื่อผู้มาประชุมมาครบองค์ประชุมแล้ว ' + d.chair +
      ' กล่าวเปิดการประชุมและเริ่มดำเนินการประชุมตามระเบียบวาระ ดังต่อไปนี้')],
      { ind: { firstLine: 1.25 }, align: 'thaiDistribute' }));

    if (opt.thaiTest) {
      out.push(para([run('บรรทัดทดสอบอักขระไทย (ส่วนนี้มีไว้ทดสอบเท่านั้น)', { b: true, u: true })], { before: 12, keepNext: true }));
      [
        'สระบน-ล่างและวรรณยุกต์ซ้อน: น้ำ ผู้ใหญ่ ปั๊ม กี่ คั่น ที่ปรึกษา ญาติ ฐาน ฎีกา ปฏิบัติ ฏ ฐ ญ',
        'การันต์และอักษรพิเศษ: ศาสตร์ จันทร์ พระพุทธเจ้า ฤทธิ์ ฤๅษี ฦๅ ฯลฯ ฯ ๆ',
        'เลขไทย ๐๑๒๓๔๕๖๗๘๙ / เลขอารบิก 0123456789 / ภาษาอังกฤษปน Zoom Meetings',
      ].forEach(function (line) { out.push(para(line)); });
      out.push(para([run('ตัวหนา: ผู้ใหญ่ปั๊มน้ำ ', { b: true }), run('ขีดเส้นใต้: ที่ปรึกษากี่ท่าน', { u: true })]));
      out.push(para([run('ย่อหน้ายาวไม่มีเว้นวรรค (ทดสอบการตัดคำท้ายบรรทัด): ' +
        'การประชุมครั้งนี้มีวัตถุประสงค์เพื่อพิจารณาหลักสูตรใหม่และหลักสูตรปรับปรุงของคณะต่างๆภายในมหาวิทยาลัย' +
        'ซึ่งผ่านการพิจารณาจากคณะกรรมการวิชาการแล้วโดยไม่มีการเว้นวรรคตลอดทั้งย่อหน้าเพื่อทดสอบว่าโปรแกรมตัดคำไทยได้ถูกต้องที่ขอบบรรทัดหรือไม่')],
        { ind: { firstLine: 1.25 }, align: 'thaiDistribute' }));
    }
    return out.join('');
  }

  function sectPr() {
    return '<w:sectPr><w:footerReference w:type="default" r:id="rId3"/>' +
      '<w:pgSz w:w="' + PAGE.width + '" w:h="' + PAGE.height + '"/>' +
      '<w:pgMar w:top="' + tw(PAGE.top) + '" w:right="' + tw(PAGE.right) + '" w:bottom="' + tw(PAGE.bottom) +
      '" w:left="' + tw(PAGE.left) + '" w:header="' + tw(PAGE.header) + '" w:footer="' + tw(PAGE.footer) +
      '" w:gutter="0"/></w:sectPr>';
  }

  function stylesXml() {
    var sz = PAGE.fontSize * 2;
    return HEAD + '<w:styles xmlns:w="' + NS_W + '">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="' + FONT + '" w:hAnsi="' + FONT + '" w:eastAsia="' + FONT + '" w:cs="' + FONT + '"/>' +
      '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>' +
      '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="th-TH"/>' +
      '</w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
      '</w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>' +
      '</w:styles>';
  }

  function settingsXml() {
    return HEAD + '<w:settings xmlns:w="' + NS_W + '">' +
      '<w:defaultTabStop w:val="720"/>' +
      '<w:characterSpacingControl w:val="doNotCompress"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
      '<w:themeFontLang w:val="en-US" w:bidi="th-TH"/>' +
      '</w:settings>';
  }

  function footerXml() {
    var s = PAGE.footerSize;
    return HEAD + '<w:ftr xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '"><w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>' +
      field('PAGE', '1', s) + run('/', { size: s }) + field('NUMPAGES', '1', s) +
      '</w:p></w:ftr>';
  }

  function coreXml(title) {
    var now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
    return HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(title) + '</dc:title><dc:creator>meeting-system (phase 0)</dc:creator>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>' +
      '</cp:coreProperties>';
  }

  var CONTENT_TYPES = HEAD +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '</Types>';

  var ROOT_RELS = HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '</Relationships>';

  var DOC_RELS = HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
    '</Relationships>';

  /** สร้าง .docx คืนค่าเป็น Blob */
  function buildReportHeader(data, opt) {
    opt = opt || {};
    var doc = HEAD + '<w:document xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '"><w:body>' +
      buildBody(data, opt) + sectPr() + '</w:body></w:document>';
    return global.MiniZip.createZip([
      { name: '[Content_Types].xml', data: CONTENT_TYPES },
      { name: '_rels/.rels', data: ROOT_RELS },
      { name: 'docProps/core.xml', data: coreXml(data.header[0]) },
      { name: 'word/document.xml', data: doc },
      { name: 'word/_rels/document.xml.rels', data: DOC_RELS },
      { name: 'word/styles.xml', data: stylesXml() },
      { name: 'word/settings.xml', data: settingsXml() },
      { name: 'word/footer1.xml', data: footerXml() }
    ], DOCX_MIME);
  }

  global.DocxPhase0 = { buildReportHeader: buildReportHeader, FONT: FONT, PAGE: PAGE };
})(window);
