"""Build the cold / hot Word templates from Raphael's 'Commissioning_Hyperstrong.docx'.

Each template keeps the original document (cover, owner table, basic information, notes,
tools, abbreviations, evaluation, notes, appendix) and gets tokens/markers that the app fills in:
  {{f.<field>}}  text of a form field          {{units.bess}} / {{units.pcs}}  unit name: SN lines
  {{sig.sN.name}} / {{sig.sN.date}} / {{sigimg.sN}}  sign-off
  marker paragraphs  {{CHECKLIST}}  {{UNITS}}  {{IPS}}  {{TOOLS}}
"""
import copy, os, re, shutil, subprocess, sys, zipfile
from lxml import etree

SRC = sys.argv[1]
OUT = sys.argv[2]
WORK = os.path.join(os.path.dirname(OUT), '_tplwork')
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'w': W, 'r': R, 'v': 'urn:schemas-microsoft-com:vml', 'o': 'urn:schemas-microsoft-com:office:office'}
q = lambda t: '{%s}%s' % (W, t.split(':')[1]) if t.startswith('w:') else t

shutil.rmtree(WORK, ignore_errors=True)
os.makedirs(WORK)
with zipfile.ZipFile(SRC) as z:
    z.extractall(WORK)

doc_path = f'{WORK}/word/document.xml'
rels_path = f'{WORK}/word/_rels/document.xml.rels'
tree = etree.parse(doc_path)
root = tree.getroot()
body = root.find('w:body', NS)
rels = etree.parse(rels_path)
relmap = {r.get('Id'): r for r in rels.getroot()}

# ---- 1. Replace the 4 Visio OLE objects (EMF previews) with plain PNG pictures -------------
def emf_to_png(emf):
    out = os.path.dirname(emf)
    subprocess.run(['soffice', '--headless', '--convert-to', 'png', '--outdir', out, emf], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
    return emf[:-4] + '.png'

INLINE = '''<w:drawing xmlns:w="{W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="{R}"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{id}" name="Picture {id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="{id}" name="Picture {id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'''
pid = 9000
for obj in body.iter(q('w:object')):
    shape = obj.find('.//v:shape', NS)
    img = obj.find('.//v:imagedata', NS)
    if shape is None or img is None:
        continue
    rid = img.get('{%s}id' % R)
    style = shape.get('style', '')
    wpt = float(re.search(r'width:([\d.]+)pt', style).group(1))
    hpt = float(re.search(r'height:([\d.]+)pt', style).group(1))
    target = relmap[rid].get('Target')
    png = emf_to_png(f'{WORK}/word/{target}')
    relmap[rid].set('Target', target[:-4] + '.png')
    pid += 1
    dr = etree.fromstring(INLINE.format(W=W, R=R, cx=int(wpt * 12700), cy=int(hpt * 12700), id=pid, rid=rid))
    obj.getparent().replace(obj, dr)
# drop the OLE relationships (no longer referenced)
for r in list(rels.getroot()):
    if r.get('Type').endswith('/oleObject'):
        rels.getroot().remove(r)

kids = list(body)
B = lambda i: kids[i]

# ---- helpers ------------------------------------------------------------------------------
def tok_run(token, rpr=None):
    r = etree.SubElement(etree.Element(q('w:p')), q('w:r'))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, q('w:t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = token
    return r

def cell_set(tc, token, like=None, bold=None):
    """Put a single run holding `token` in the cell's first paragraph (keeps paragraph format)."""
    ps = tc.findall('w:p', NS)
    p = ps[0]
    for extra in ps[1:]:
        tc.remove(extra)
    src_rpr = None
    for cand in ([like] if like is not None else []) + [tc]:
        rr = cand.find('.//w:r/w:rPr', NS)
        if rr is not None:
            src_rpr = copy.deepcopy(rr); break
    for r in p.findall('w:r', NS) + p.findall('w:sdt', NS) + p.findall('w:hyperlink', NS):
        p.remove(r)
    if src_rpr is not None and bold is False:
        for b in src_rpr.findall('w:b', NS) + src_rpr.findall('w:bCs', NS):
            src_rpr.remove(b)
    p.append(tok_run(token, src_rpr))

def tcs(tr):
    return tr.findall('w:tc', NS)

def marker(name):
    p = etree.Element(q('w:p'))
    p.append(tok_run('{{%s}}' % name))
    return p

# ---- 2. Basic information table -------------------------------------------------------------
info = B(18)
rows = info.findall('w:tr', NS)
def val(tr, idx, token, label_idx=None):
    cells = tcs(tr)
    like = cells[label_idx] if label_idx is not None else None
    cell_set(cells[idx], token, like=like, bold=False)
def lab(tr, idx, text, like_tr, like_idx):
    cell_set(tcs(tr)[idx], text, like=tcs(like_tr)[like_idx])

print('row cell counts', [len(tcs(r)) for r in rows])
val(rows[1], 1, '{{f.projectName}}', 0); val(rows[1], 3, '{{f.customer}}', 2)
val(rows[2], 1, '{{f.epc}}', 0); val(rows[2], 3, '{{f.site}}', 2)
val(rows[3], len(tcs(rows[3])) - 1 if len(tcs(rows[3])) > 1 else 1, '{{f.pmContact}}', 0)
val(rows[4], len(tcs(rows[4])) - 1 if len(tcs(rows[4])) > 1 else 1, '{{f.location}}', 0)
val(rows[7], 1, '{{f.batteryType}}', 0); val(rows[7], 3, '{{units.bess}}', 2)
val(rows[10], 1, '{{f.mwType}}', 0); val(rows[10], 3, '{{f.transformerSN}}', 2)

# blank row under MW Unit: MV SN + PCS units
mvrow = copy.deepcopy(rows[10])
lab(mvrow, 0, 'MV SN', rows[10], 0); val(mvrow, 1, '{{f.mvSN}}', 0)
lab(mvrow, 2, 'PCS SN', rows[10], 2); val(mvrow, 3, '{{units.pcs}}', 2)
rows[10].addnext(mvrow)
val(rows[13], 1, '{{f.rmuType}}', 0); val(rows[13], 3, '{{f.rmuSN}}', 2)
# the tall empty rows under Battery Unit / MW Unit are the photo areas: BESS / PCS SN plate photos
for ri, tk in ((8, 'photos.bess'), (11, 'photos.pcs')):
    p0 = tcs(rows[ri])[0].find('w:p', NS)
    p0.append(tok_run('{{%s}}' % tk))
# Commissioning visit block (copies of the MV/RMU header + label rows), before the purpose row
hdr, lbl = rows[12], rows[13]
new = []
h = copy.deepcopy(hdr)
for c in tcs(h):
    cell_set(c, 'Commissioning visit')
new.append(h)
for a, ta, b2, tb in [('HyperStrong engineer', 'engineer', 'Supervising person', 'supervisor'),
                      ('Service partner company', 'partner', 'Date', 'date'),
                      ('Start time', 'arrival', 'End time', 'departure')]:
    r = copy.deepcopy(lbl)
    cs = tcs(r)
    cell_set(cs[0], a, like=tcs(lbl)[0]); cell_set(cs[1], '{{f.%s}}' % ta, like=tcs(lbl)[0], bold=False)
    cell_set(cs[2], b2, like=tcs(lbl)[2]); cell_set(cs[3], '{{f.%s}}' % tb, like=tcs(lbl)[2], bold=False)
    new.append(r)
purpose = rows[15]
for r in new:
    purpose.addprevious(r)

# ---- 3. Evaluation tables + comments ------------------------------------------------------------
def eval_tokens(tbl):
    rws = tbl.findall('w:tr', NS)
    c0 = tcs(rws[0])
    uniq = []
    for c in c0:
        if c not in uniq: uniq.append(c)
    for c, tokname in zip(uniq, ['evalCustomer', 'evalCustomer', 'evalHS', 'evalHS']):
        txt = ''.join(c.xpath('.//w:t/text()', namespaces=NS))
        if 'costumer' in txt.lower() or 'customer' in txt.lower(): tk = 'evalCustomer'
        elif 'hyperstrong' in txt.lower(): tk = 'evalHS'
        else: continue
        ps = c.findall('w:p', NS)
        np_ = copy.deepcopy(ps[0])
        for r in np_.findall('w:r', NS): np_.remove(r)
        np_.append(tok_run('{{f.%s}}' % tk))
        ps[-1].addnext(np_)
    # signature cells: HyperStrong Engineer, Service Partner, (Service Partner), Customer
    cells = tcs(rws[1])
    mapping = {i: 's%d' % i for i in range(min(3, len(cells)))}
    for ci, sk in mapping.items():
        ps = cells[ci].findall('w:p', NS)
        lines = [i for i, p in enumerate(ps) if '____' in ''.join(p.xpath('.//w:t/text()', namespaces=NS))]
        if len(lines) >= 3:
            for li, tk in zip(lines[:2], ['name', 'date']):
                p = ps[li]
                rpr = p.find('.//w:r/w:rPr', NS)
                for r in p.findall('w:r', NS): p.remove(r)
                p.append(tok_run('{{sig.%s.%s}}' % (sk, tk), rpr))
            p = ps[lines[2]]
            rpr = p.find('.//w:r/w:rPr', NS)
            for r in p.findall('w:r', NS): p.remove(r)
            p.append(tok_run('Signature: ', rpr))
            p.append(tok_run('{{sigimg.%s}}' % sk, rpr))

eval_cold, eval_hot = B(94), B(141)
eval_tokens(eval_cold); eval_tokens(eval_hot)
comments_p = B(143)
for r in comments_p.findall('w:r', NS): comments_p.remove(r)
comments_p.append(tok_run('{{f.remarks}}'))


# ---- 4. Write the two variants --------------------------------------------------------------
original = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
rels_bytes = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

def variant(name, drop_ranges, markers):
    t2 = etree.fromstring(original)
    b2 = t2.find('w:body', NS)
    k2 = list(b2)
    for idx, mk in markers:          # insert marker before element idx (original numbering)
        k2[idx].addprevious(marker(mk))
    for a, z in drop_ranges:
        for i in range(a, z + 1):
            b2.remove(k2[i])
    open(doc_path, 'wb').write(etree.tostring(t2, xml_declaration=True, encoding='UTF-8', standalone=True))
    open(rels_path, 'wb').write(rels_bytes)
    # content types: png already declared
    out = OUT.replace('{name}', name)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for dp, _, fs in os.walk(WORK):
            for f in fs:
                full = os.path.join(dp, f); arc = os.path.relpath(full, WORK)
                if arc.startswith('word/embeddings') or arc.endswith('.emf'):
                    continue
                z.write(full, arc)
    print('wrote', out)

# index map (after the UNITS marker insert, the original kids list is unchanged in our copy: markers are new nodes)
# cold: keep H1 54 + intro 55, checklist marker replaces 56..91, drop the whole hot part 101..141
variant('cold', [(56, 92), (101, 141)], [(20, 'UNITS'), (56, 'CHECKLIST')])
# hot: drop the cold part 54..100, IP table 102 -> IPS marker (+TOOLS), checklist 104..134 -> CHECKLIST
variant('hot', [(54, 100), (102, 102), (103, 134)], [(20, 'UNITS'), (102, 'IPS'), (102, 'TOOLS'), (104, 'CHECKLIST')])
