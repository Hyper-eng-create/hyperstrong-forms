/* HyperStrong commissioning report in the layout of "Commissioning_Hyperstrong.docx".
   The Word file is the original document (template in assets/) filled with the report;
   the PDF is that Word file drawn on screen (docx-preview) and printed page by page. */
(function(){
  if (!window.CFG || !window.CFG.docxTpl) return;
  const C = window.CFG;
  const loadScript = src => new Promise((res, rej) => { if (document.querySelector(`script[src="${src}"]`)) return res(); const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => rej(new Error("Couldn't load " + src)); document.head.appendChild(s); });
  const XE = s => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const TX = s => XE(s).replace(/\r?\n/g, '</w:t><w:br/><w:t xml:space="preserve">');
  const RUN = (text, o = {}) => `<w:r><w:rPr>${o.b ? "<w:b/><w:bCs/>" : ""}${o.color ? `<w:color w:val="${o.color}"/>` : ""}${o.sz ? `<w:sz w:val="${o.sz}"/><w:szCs w:val="${o.sz}"/>` : ""}</w:rPr><w:t xml:space="preserve">${TX(text)}</w:t></w:r>`;
  const PP = (inner, o = {}) => `<w:p><w:pPr>${o.num ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>` : ""}<w:bidi w:val="0"/>${o.ind ? `<w:ind w:left="${o.ind}"/>` : ""}${o.after != null ? `<w:spacing w:after="${o.after}"/>` : ""}${o.b ? "<w:rPr><w:b/><w:bCs/></w:rPr>" : ""}</w:pPr>${inner}</w:p>`;
  const HEAD = (text) => PP(RUN(text, {b:true}), {b:true});
  const FILL = '<w:shd w:val="clear" w:color="auto" w:fill="A5C9EB" w:themeFill="text2" w:themeFillTint="40"/>';
  const TC = (w, inner, head) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${head ? FILL : ""}</w:tcPr>${inner || PP("", {ind:720})}</w:tc>`;
  const TBL = (widths, headers, rows) => `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="720" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>
    <w:tblGrid>${widths.map(w => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>
    <w:tr><w:trPr><w:tblHeader/></w:trPr>${headers.map((h, i) => TC(widths[i], PP(RUN(h), {ind:720}), true)).join("")}</w:tr>${rows.join("")}</w:tbl>${PP("")}`;
  const STAT = {pass:["PASS", "1E7B34"], fail:["FAIL", "C00000"], na:["N/A", "595959"]};
  let sdtId = 1700000000;
  const CHECK = on => `<w:sdt><w:sdtPr><w:rPr><w:b/><w:bCs/></w:rPr><w:id w:val="${sdtId++}"/><w14:checkbox><w14:checked w14:val="${on ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/><w:b/><w:bCs/></w:rPr><w:t>${on ? "☒" : "☐"}</w:t></w:r></w:sdtContent></w:sdt>`;

  // ---- pictures ------------------------------------------------------------------------------
  function picturesFor(zip){
    let n = 0; const rels = []; let docPr = 30000;
    async function add(src, maxW, maxH){
      if (!src) return "";
      const im = new Image(); await new Promise(res => { im.onload = res; im.onerror = res; im.src = src; });
      const w0 = im.naturalWidth || 1, h0 = im.naturalHeight || 1; const sc = Math.min(1, 1600 / Math.max(w0, h0));
      const c = document.createElement("canvas"); c.width = Math.round(w0 * sc); c.height = Math.round(h0 * sc);
      const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(im, 0, 0, c.width, c.height);
      const data = c.toDataURL("image/jpeg", 0.85).split(",")[1];
      n++; const name = `hsapp_${n}.jpeg`, rid = `rIdHsApp${n}`;
      zip.file("word/media/" + name, data, {base64:true});
      rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`);
      const k = Math.min(maxW / w0, maxH / h0); const cx = Math.round(w0 * k * 914400), cy = Math.round(h0 * k * 914400); docPr++;
      return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPr}" name="Photo ${docPr}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${docPr}" name="Photo ${docPr}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
    }
    return {add, rels};
  }

  async function fillDocx(r, imgs, opt = {}){
    await loadScript(C.libs && C.libs.jszip || "lib/jszip.min.js");
    const tpl = await (await fetch(C.docxTpl)).arrayBuffer();
    const zip = await JSZip.loadAsync(tpl);
    let x = await zip.file("word/document.xml").async("string");
    const pics = picturesFor(zip);
    const f = r.f || {};
    // preset table texts may have been created in Hebrew: the report is in English like the original document
    const REV = {}; try { Object.entries(window.HE || {}).forEach(([en, he]) => { if (typeof he === "string" && !(he in REV)) REV[he] = en; }); } catch(e){}
    const EN = t => (t && REV[t]) || t;
    const photoRuns = async (list, w, h) => { let o = ""; for (const p of (list || [])) { const d = imgs[p]; if (d){ const dr = await pics.add(d, w, h); if (dr) o += `<w:r>${dr}</w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`; } } return o; };
    const unitLines = k => (f[k] || []).filter(u => u && (u.name || u.sn)).map(u => `${EN(u.name) || ""}${u.sn ? ": " + u.sn : ""}`).join("\n");

    // simple tokens
    x = x.replace(/\{\{f\.(\w+)\}\}/g, (m, k) => TX(f[k] ?? ""));
    x = x.replace(/\{\{units\.bess\}\}/g, TX(unitLines("bessUnits"))).replace(/\{\{units\.pcs\}\}/g, TX(unitLines("pcsUnits")));
    x = x.replace(/\{\{sig\.(s\d)\.(name|date)\}\}/g, (m, s, k) => TX((r.signoff?.[s] || {})[k] || "____________________"));
    const sigs = {};
    for (const s of ["s0", "s1", "s2"]){ const so = r.signoff?.[s]; sigs[s] = so && so.sig ? await pics.add(so.sig, 1.6, 0.6) : ""; }
    x = x.replace(/<w:t xml:space="preserve">\{\{sigimg\.(s\d)\}\}<\/w:t>/g, (m, s) => sigs[s] || '<w:t>____________________</w:t>');

    // BESS / PCS units
    const units = "";   // unit names / SN go in the Basic information table, SN plate photos in its photo rows
    const unitPhotos = async k => { let o = ""; for (const u of (f[k] || [])) o += await photoRuns(u && u.snPhoto, 2.2, 1.6); return o; };
    x = x.replace('<w:r><w:t xml:space="preserve">{{photos.bess}}</w:t></w:r>', await unitPhotos("bessUnits"));
    x = x.replace('<w:r><w:t xml:space="preserve">{{photos.pcs}}</w:t></w:r>', await unitPhotos("pcsUnits"));
    // IP addresses (hot)
    const ipRows = (f.ips || []).filter(u => u && (u.equipment || u.ip || u.note)).map(u => `<w:tr>${TC(3953, PP(RUN(EN(u.equipment) || ""), {ind:720}))}${TC(3686, PP(RUN((u.ip || "") + (u.note ? `  (${u.note})` : "")), {ind:720}))}</w:tr>`);
    let ips = ipRows.length ? TBL([3953, 3686], ["Equipment", "IP address "], ipRows) : "";
    const shots = (f.ipShots || []).filter(g => g && imgs[g.p]);
    if (shots.length){ let o = ""; for (const g of shots){ const dr = await pics.add(imgs[g.p], 2.6, 2.2); o += PP(`<w:r>${dr}</w:r>`, {ind:720, after:0}) + (g.cap ? PP(RUN(g.cap, {sz:18}), {ind:720}) : ""); }
      ips += HEAD("IP configuration screenshots (IP, gateway, NTP and subnet)") + o + PP(""); }
    const toolRows = (f.tools || []).filter(u => u && u.name).map(u => `<w:tr>${TC(7639, PP(RUN(EN(u.name)), {ind:0}))}</w:tr>`);
    const tools = toolRows.length ? HEAD("Test equipment") + TBL([7639], ["Tool"], toolRows) : "";

    // checklist
    let cl = "", n = 0;
    for (const sec of C.sections){
      n++; cl += HEAD(`${n}. ${sec.t}${sec.sub ? ` (${sec.sub})` : ""}`);
      for (const g of sec.groups){
        if (g.t) cl += HEAD(g.t);
        const rows = [];
        for (const it of g.items){
          const v = r.results?.[it.id] || {}; const st = STAT[v.s];
          const status = PP(CHECK(v.s === "pass") + (st ? RUN(" " + st[0], {b:true, color:st[1]}) : ""), {ind:720, b:true});
          let com = "";
          if (v.v) com += PP(RUN(`Value: ${v.v} ${it.val?.unit || ""}`.trim(), {b:true}), {after:0});
          if (v.c) com += PP(RUN(v.c), {after:0});
          const ph = await photoRuns(v.photos, 1.25, 1.0); if (ph) com += PP(ph, {after:0});
          rows.push(`<w:tr><w:trPr><w:cantSplit/></w:trPr>${TC(3484, PP(RUN(it.t + (it.optional ? " (optional)" : "")), {num:true}))}${TC(2046, status)}${TC(2046, com || PP("", {ind:720}))}</w:tr>`);
        }
        cl += TBL([3484, 2046, 2046], ["Item", "Status", "Comments"], rows);
      }
    }
    const MARK = k => new RegExp(`<w:p><w:r><w:t xml:space="preserve">\\{\\{${k}\\}\\}</w:t></w:r></w:p>`);
    x = x.replace(MARK("UNITS"), () => units).replace(MARK("IPS"), () => ips).replace(MARK("TOOLS"), () => tools).replace(MARK("CHECKLIST"), () => cl);
    zip.file("word/document.xml", x);
    let rels = await zip.file("word/_rels/document.xml.rels").async("string");
    rels = rels.replace("</Relationships>", pics.rels.join("") + "</Relationships>");
    zip.file("word/_rels/document.xml.rels", rels);
    if (!/Extension="jpeg"/i.test(await zip.file("[Content_Types].xml").async("string"))){
      let ct = await zip.file("[Content_Types].xml").async("string"); ct = ct.replace("<Types ", '<Types ').replace(/(<Types[^>]*>)/, '$1<Default Extension="jpeg" ContentType="image/jpeg"/>'); zip.file("[Content_Types].xml", ct); }
    if (opt.pdf && zip.file("word/numbering.xml")){   // on screen: plain bullets instead of Symbol-font glyphs
      let nb = await zip.file("word/numbering.xml").async("string");
      nb = nb.replace(/w:lvlText w:val="[\uF000-\uF0FF]"/g, 'w:lvlText w:val="•"').replace(/w:(ascii|hAnsi|cs)="(Symbol|Wingdings)"/g, 'w:$1="Arial"');
      zip.file("word/numbering.xml", nb); }
    return await zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
  }

  // ---- PDF: draw the Word file and print it page by page ---------------------------------
  async function docxToPdf(blob){
    await loadScript(C.libs && C.libs.jszip || "lib/jszip.min.js");
    const prevDocx = window.docx; window.docx = undefined;
    await loadScript(C.libs && C.libs.docxPreview || "lib/docx-preview.min.js");
    const DP = window.docx; window.docxPreview = DP; window.docx = prevDocx;
    await loadScript((C.libs && C.libs.h2c) || "lib/html2canvas.min.js"); await loadScript((C.libs && C.libs.jspdf) || "lib/jspdf.umd.min.js");
    const host = document.createElement("div");
    host.style.cssText = "position:fixed;left:-10000px;top:0;width:900px;background:#fff;z-index:-1;direction:ltr;text-align:left";
    host.className = "notr"; host.setAttribute("dir", "ltr"); host.setAttribute("lang", "en"); host.setAttribute("translate", "no");
    const fontCss = document.createElement("style");
    fontCss.textContent = '.hsdocx section.docx, .hsdocx section.docx *{font-family:Aptos,"Segoe UI",Calibri,Carlito,"Helvetica Neue",Arial,sans-serif !important} .hsdocx section.docx p, .hsdocx section.docx span{word-break:normal !important;overflow-wrap:normal !important;hyphens:none !important}';
    host.classList.add("hsdocx"); document.head.appendChild(fontCss);
    document.body.appendChild(host);
    try {
      await DP.renderAsync(blob, host, null, {inWrapper:false, breakPages:true, ignoreLastRenderedPageBreak:true, renderHeaders:true, renderFooters:true, ignoreWidth:false, ignoreHeight:false, experimental:true, useBase64URL:true});
      await new Promise(res => setTimeout(res, 300));
      await Promise.all([...host.querySelectorAll("img")].map(im => im.complete ? 0 : new Promise(r => { im.onload = im.onerror = r; })));
      const pages = paginate(host);
      const { jsPDF } = window.jspdf; const pdf = new jsPDF({unit:"pt", format:"a4", compress:true});
      for (let i = 0; i < pages.length; i++){
        const cv = await html2canvas(pages[i], {scale:2, backgroundColor:"#ffffff", useCORS:true, logging:false});
        if (i) pdf.addPage();
        pdf.addImage(cv.toDataURL("image/jpeg", 0.88), "JPEG", 0, 0, 595.28, 841.89);
      }
      return pdf.output("blob");
    } finally { host.remove(); fontCss.remove(); }
  }
  // Split each rendered Word section into A4 pages at paragraph / table-row boundaries, repeating the header.
  function paginate(host){
    const out = [];
    const sections = [...host.querySelectorAll("section")].filter(s => s.querySelector("article"));
    for (const sec of sections){
      const H = sec.offsetWidth * 297 / 210; const art = sec.querySelector("article");
      const header = sec.querySelector("header"), footer = sec.querySelector("footer");
      const cs = getComputedStyle(sec); const padT = parseFloat(cs.paddingTop) || 0, padB = parseFloat(cs.paddingBottom) || 0;
      const avail = H - padT - padB - (header ? header.offsetHeight : 0) - (footer ? footer.offsetHeight : 0) - 4;
      const blocks = [...art.children];
      const newPage = () => { const pg = sec.cloneNode(false); pg.style.minHeight = H + "px"; pg.style.height = H + "px"; pg.style.overflow = "hidden"; pg.style.boxSizing = "border-box";
        if (header) pg.appendChild(header.cloneNode(true)); const a = art.cloneNode(false); pg.appendChild(a); if (footer) pg.appendChild(footer.cloneNode(true)); sec.parentNode.insertBefore(pg, sec); out.push(pg); return a; };
      let a = newPage(); let used = 0;
      const fullH = el => el.offsetHeight + parseFloat(getComputedStyle(el).marginTop || 0) + parseFloat(getComputedStyle(el).marginBottom || 0);
      const isHeadRow = tr => tr && /\b(Item|Equipment|Tool|Abbreviation|No)\b/.test((tr.textContent || "").trim().slice(0, 40)) && /Status|IP address|Serial|Full name|Quantity/.test(tr.textContent || "");
      for (let bi = 0; bi < blocks.length; bi++){
        const bl = blocks[bi]; const h = fullH(bl);
        if (bl.tagName === "TABLE"){
          if (used + h <= avail){ a.appendChild(bl.cloneNode(true)); used += h; continue; }
          if (h <= avail){ a = newPage(); used = 0; a.appendChild(bl.cloneNode(true)); used += h; continue; }
          const rows = [...bl.querySelectorAll(":scope > tbody > tr, :scope > tr")]; const head = isHeadRow(rows[0]) ? rows[0] : null;
          if (used > 0 && used + (rows[0] ? rows[0].offsetHeight : 0) + (rows[1] ? rows[1].offsetHeight : 0) > avail){ a = newPage(); used = 0; }
          let t = bl.cloneNode(false); let tb = document.createElement("tbody"); t.appendChild(tb); a.appendChild(t);
          for (const tr of rows){ const rh = tr.offsetHeight;
            if (used + rh > avail && tb.children.length){ a = newPage(); used = 0; t = bl.cloneNode(false); tb = document.createElement("tbody"); t.appendChild(tb); a.appendChild(t); if (head && tr !== head){ tb.appendChild(head.cloneNode(true)); used += head.offsetHeight; } }
            tb.appendChild(tr.cloneNode(true)); used += rh; }
          continue; }
        if (!(bl.textContent || "").trim() && !bl.querySelector("img,svg,canvas")){ if (used + h <= avail){ a.appendChild(bl.cloneNode(true)); used += h; } continue; }
        // keep a heading with the start of the table that follows it
        let nx = blocks[bi + 1]; let need = h;
        if (nx && nx.tagName !== "TABLE" && (bl.textContent || "").trim().length < 90 && (nx.textContent || "").trim().length < 90 && blocks[bi + 2] && blocks[bi + 2].tagName === "TABLE"){ need += fullH(nx); nx = blocks[bi + 2]; }
        if (nx && nx.tagName === "TABLE" && (bl.textContent || "").trim().length < 90){ const rws = nx.querySelectorAll("tr"); const th = fullH(nx); need += th <= avail ? th : (rws[0] ? rws[0].offsetHeight : 0) + (rws[1] ? rws[1].offsetHeight : 0) + 6; }
        if (used + need > avail && used > 0){ a = newPage(); used = 0; }
        a.appendChild(bl.cloneNode(true)); used += h;
      }
      sec.style.display = "none";
    }
    return out;
  }

  window.HS_EXPORT = {word: (r, imgs) => fillDocx(r, imgs), pdf: async (r, imgs) => docxToPdf(await fillDocx(r, imgs, {pdf:true}))};
})();
