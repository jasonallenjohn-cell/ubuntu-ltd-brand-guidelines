"""Co-brand Trillium client material with a realtor's brand: Trillium leads, the realtor presents.

    python3 cobrand.py deck     <trillium deck root> <profile.json> <out root> [--drop stress,next] [--rewrites deal.json]
    python3 cobrand.py onepager <trillium one-pager.html> <profile.json> <out.html>
    python3 cobrand.py check    <profile.json>

<trillium deck root> is a Slides deck folder (holds project/deck.json and project/slides/*.html) in the Trillium
deck markup: each content slide has the "Trillium Company" header cell, the cover has the "Investment summary" line.
The script never writes into the source; it writes a new deck root ready to publish as its own Slides artifact.

What it does
- Header: "Trillium Company" becomes "Trillium Company x <brand>" on every slide.
- Cover: a "Presented by" block in the realtor's dark colour with an accent rule.
- Closing: a "Let's connect" slide in the realtor's colours (name, brokerage, contact, three next steps, Trillium mark).
- Client-safe: table rows whose first cell names an internal line (commission, bonus, refinance surplus, fee budget,
  profit) are removed; slides named in --drop are left out; any other blocked phrase left anywhere stops the build
  with the slide and the phrase, so a person rewrites that line (nothing internal ships by accident).
- Rewrites: --rewrites names a JSON file {slide id: [[old text, new text], ...], "*": [...]} applied before the
  check, for the deal-specific lines (a stat or a speaker note) that must read differently for clients. A slide's
  speaker notes can be replaced whole with ["<aside>", "new notes"].
- Compliance: profile check flags placeholder phone numbers (555-01xx), a missing brokerage legal name or
  registration line, and prints the profile's to_confirm list. Missing items never print as blanks on the page.
"""
import html, json, os, re, shutil, sys

BLOCKED = [r'commission', r'\bbonus\b', r'refinance surplus', r'paid to trillium', r'fee budget', r'\bprofit\b',
           r'trillium fee', r'development management fee', r'lender fee', r'broker fee', r'surplus after tpi']
ROW_BLOCK = re.compile(r'<tr[^>]*>\s*<td[^>]*>\s*(?:TPO commission|Refinance surplus[^<]*|Trillium bonus[^<]*|'
                       r'Fee[^<]*|Real estate commission[^<]*|Profit[^<]*|Commission[^<]*)</td>.*?</tr>\n?', re.S | re.I)
DEFAULT_DROP = ['stress', 'next', 'refi']
DISCLAIMER = ('Presented by {name}, {brand}, {brokerage_line}. {first} introduces property owners to Trillium; the '
              'Trillium Property Investment (TPI) and Trillium Property Ownership (TPO) programs are offered and '
              'administered by Trillium only. Figures are projections; this is not an offer to sell securities.')


def load(p):
    pr = json.load(open(p))
    c = pr.setdefault('colors', {})
    for k, v in {'dark': '#15171A', 'accent': '#C9A36A', 'text': '#F3F1EC', 'soft': '#C9C7C1', 'muted': '#8E8D88',
                 'line': '#373B41'}.items():
        c.setdefault(k, v)
    pr.setdefault('font', {'family': 'Inter', 'google': 'Inter:wght@400;600;700'})
    pr.setdefault('presenter_label', 'Your advisor on this property')
    pr.setdefault('brand_mark', '')
    return pr


def esc(s):
    return html.escape(s or '', quote=True)


def check(pr):
    warn = []
    if pr.get('phone') and re.search(r'555[-\s.]?01\d\d', pr['phone']):
        warn.append(f"phone {pr['phone']} is a 555-01xx placeholder: remove it or confirm the real number")
    if not pr.get('brokerage_legal'):
        warn.append('brokerage_legal missing: RECO wants the registered brokerage name on client material')
    if not pr.get('registration'):
        warn.append('registration line missing (title / brokerage statement)')
    if not pr.get('email') and not pr.get('phone'):
        warn.append('no email or phone: the contact block will be thin')
    for t in pr.get('to_confirm', []):
        warn.append('to confirm: ' + t)
    return warn


def brokerage_line(pr):
    return pr.get('brokerage_legal') or pr.get('brokerage') or ''


def contacts(pr):
    out = [pr.get('email'), pr.get('phone') if pr.get('phone') and not re.search(r'555[-\s.]?01\d\d', pr['phone']) else None,
           pr.get('social'), pr.get('website')]
    return [x for x in out if x]


def wordmark(pr, ff):
    c = pr['colors']
    if pr.get('logo_blob'):
        return f'<img src="/_blob/{esc(pr["logo_blob"])}" alt="{esc(pr["brand"])} logo" style="height:30px;width:auto;object-fit:contain">'
    mark = f'<span style="color:{c["accent"]}">{esc(pr["brand_mark"])}</span>' if pr.get('brand_mark') else ''
    return f'<span style="{ff}">{esc(pr["brand"])}{mark}</span>'


def blocked_hits(text):
    t = re.sub(r'<[^>]+>', ' ', text).lower()
    return [b for b in BLOCKED if re.search(b, t)]


# ------------------------------------------------------------------ deck
def deck(src_root, prof, out_root, drop, rewrites=None):
    pr = load(prof)
    c = pr['colors']
    ff = f"font-family:'{pr['font']['family']}', Arial, sans-serif"
    src = os.path.join(src_root, 'project')
    dj = json.load(open(os.path.join(src, 'deck.json')))
    order = [s for s in dj['order'] if s not in drop]
    os.makedirs(os.path.join(out_root, 'project', 'slides'), exist_ok=True)
    WM = wordmark(pr, ff)
    slides, problems = {}, []
    for sid in order:
        h = open(os.path.join(src, 'slides', f'{sid}.html')).read()
        for col in ('#2A453D', '#F2F1E7'):
            h = h.replace(f'<p style="font-size:24px;font-weight:700;line-height:1.25;color:{col}">Trillium Company</p>',
                          f'<p style="font-size:24px;font-weight:700;line-height:1.25;color:{col}">Trillium Company '
                          f'<span style="font-weight:300">×</span> {WM}</p>')
        for old, new in (rewrites or {}).get(sid, []) + (rewrites or {}).get('*', []):
            if old == '<aside>':
                h = re.sub(r'<aside>.*?</aside>', lambda m: f'<aside>{esc(new)}</aside>', h, flags=re.S)
            elif old in h:
                h = h.replace(old, new)
            else:
                problems.append(f'{sid}: rewrite text not found: {old[:60]}')
        h = ROW_BLOCK.sub('', h)
        if sid == dj.get('cover', order[0]) or sid == 'cover':
            h, n = re.subn(r'<p style="font-size:24px;line-height:1.25;color:#C3BD87;width:760px">(Investment summary[^<]*)</p>',
                           lambda m: f'''<div style="display:flex;flex-direction:column;gap:20px;width:760px">
<div style="display:flex;flex-direction:column;gap:6px;background:{c['dark']};border-top:2px solid {c['accent']};padding:24px 28px">
<p style="{ff};font-size:18px;font-weight:600;letter-spacing:3px;line-height:1.2;text-transform:uppercase;color:{c['muted']}">Presented by</p>
<p style="{ff};font-size:36px;font-weight:700;letter-spacing:-1px;line-height:1.1;color:{c['text']}">{esc(pr['name'])}</p>
<p style="{ff};font-size:22px;font-weight:600;letter-spacing:2px;line-height:1.2;text-transform:uppercase;color:{c['accent']}">{esc(pr['brand'])} · {esc(pr.get('brokerage'))}</p>
</div>
<p style="font-size:24px;line-height:1.25;color:#C3BD87">{m.group(1)}</p>
</div>''', h)
            if not n:
                problems.append(f'{sid}: no "Investment summary" line found for the presenter block')
            h = re.sub(r'<aside>', f'<aside>Co-branded: Trillium Company presented by {esc(pr["name"])}, {esc(pr["brand"])}, {esc(brokerage_line(pr))}. ', h, count=1)
        for hit in blocked_hits(h):
            problems.append(f'{sid}: internal wording "{hit}" left on a client slide; rewrite that line or --drop the slide')
        slides[sid] = h
    logo = re.search(r'<svg.*?</svg>', slides.get('cover', ''), re.S)
    slides['connect'] = connect_slide(pr, ff, WM, logo.group(0) if logo else '', dj.get('title', ''))
    order.append('connect')
    if problems:
        raise SystemExit('not built, fix these first:\n  ' + '\n  '.join(problems))
    for sid, h in slides.items():
        open(os.path.join(out_root, 'project', 'slides', f'{sid}.html'), 'w').write(h)
    prop = re.split(r'\s+[—-]\s+', dj.get('title', 'Trillium'))[0]
    dj['title'] = f'{prop} — Trillium × {pr["brand"]}'
    dj['order'] = order
    dj['sections'] = {k: v for k, v in dj.get('sections', {}).items() if v.get('start') in order}
    dj['sections']['zz'] = {'description': f'Next steps with {pr["name"]}.', 'start': 'connect'}
    key = pr['font']['family'].lower().replace(' ', '-')
    dj.setdefault('faces', {})[key] = {'family': pr['font']['family'],
                                       'href': f'https://fonts.googleapis.com/css2?family={pr["font"]["google"]}&display=swap'}
    dj['designSystems'] = []
    dj.update({'v': 5, 'lists': 'css', 'format': {'v': 1, 'shared': ['project/deck.css']}})
    dj['createdOnFiles'] = {'v': 1, 'at': '2026-11-15T00:00:00Z'}
    json.dump(dj, open(os.path.join(out_root, 'project', 'deck.json'), 'w'), indent=2, ensure_ascii=False)
    open(os.path.join(out_root, 'project', 'deck.css'), 'w').write('section { box-sizing: border-box; }\ntable { width: 100%; }\n')
    blobs = sorted(set(re.findall(r'/_blob/([0-9a-f]{32})', ''.join(slides.values()))))
    print(json.dumps({'out': out_root, 'order': order, 'blobs_to_copy': blobs, 'warnings': check(pr)}, indent=1))


def connect_slide(pr, ff, WM, logo, title):
    c = pr['colors']
    prop = re.split(r'\s+[—-]\s+', title)[0] if title else 'Trillium'

    def step(n, t, last=False):
        bb = f'border-bottom:1px solid {c["line"]};' if last else ''
        return (f'<div style="display:flex;flex-direction:row;gap:28px;border-top:1px solid {c["line"]};{bb}padding:24px 0px">'
                f'<p style="{ff};font-size:28px;font-weight:600;line-height:1.3;color:{c["accent"]};width:56px">{n}</p>'
                f'<p style="font-size:28px;line-height:1.35;color:{c["text"]};width:760px">{t}</p></div>')
    lines = ''.join(f'<p style="font-size:32px;line-height:1.35;color:{c["text"] if i == 0 else c["soft"]}">{esc(x)}</p>'
                    for i, x in enumerate(contacts(pr)))
    sub = ' · '.join(x for x in [pr.get('title'), pr['brand'], pr.get('brokerage')] if x)
    first = pr['name'].split()[0]
    return f'''<section id="connect" data-transition="fade" style="background:{c['dark']};color:{c['text']};font-family:'Archivo', Arial, sans-serif;padding:192px 128px 128px;display:grid;grid-template-columns:760px 872px;gap:32px">
<div style="position:absolute;left:128px;top:48px;width:1664px;height:73px;display:flex;flex-direction:row;align-items:center;justify-content:space-between;border-top:1px solid {c['accent']};border-bottom:1px solid {c['line']}">
<p style="font-size:24px;font-weight:700;line-height:1.25;color:{c['text']}">Trillium Company <span style="font-weight:300">×</span> {WM}</p>
<p style="font-size:24px;line-height:1.25;text-transform:uppercase;color:{c['soft']}">{esc(prop)} · Let’s connect</p>
</div>
<div style="display:flex;flex-direction:column;gap:32px;padding:32px 0px 0px 0px">
<p style="{ff};font-size:24px;font-weight:600;letter-spacing:3px;line-height:1.25;text-transform:uppercase;color:{c['accent']}">{esc(pr['presenter_label'])}</p>
<h2 style="{ff};font-size:88px;font-weight:700;letter-spacing:-2px;line-height:1.05;color:{c['text']}">{esc(pr['name'])}</h2>
<p style="{ff};font-size:28px;font-weight:600;letter-spacing:3px;line-height:1.25;text-transform:uppercase;color:{c['muted']}">{esc(sub)}</p>
<div style="display:flex;flex-direction:column;gap:12px;border-top:1px solid {c['line']};padding:28px 0px 0px 0px">{lines}</div>
</div>
<div style="display:flex;flex-direction:column;justify-content:space-between;padding:32px 0px 0px 0px">
<div style="display:flex;flex-direction:column;gap:0px">
{step('01', 'A free Arterial assessment of your property: what it can become and what it could earn.')}
{step('02', f'A meeting with {esc(first)} and Trillium to walk through the plan, the numbers and the timeline.')}
{step('03', 'Trillium handles design, approvals, financing, construction and property management.', True)}
</div>
<div style="display:flex;flex-direction:row;align-items:center;gap:9px">
{logo}
<div style="display:flex;flex-direction:row;align-items:baseline">
<p style="font-size:36px;font-weight:700;letter-spacing:-1px;line-height:1;color:#C3BD87">trillium</p>
<p style="font-size:36px;font-weight:300;letter-spacing:-1px;line-height:1;color:{c['text']}">company</p>
</div>
</div>
</div>
<aside>{esc(DISCLAIMER.format(name=pr['name'], brand=pr['brand'], brokerage_line=brokerage_line(pr), first=first))}</aside>
</section>
'''


# ------------------------------------------------------------------ one-pager
def onepager(src, prof, out):
    pr = load(prof)
    c = pr['colors']
    F = pr['font']['family']
    h = open(src).read()
    if 'class="pres"' in h:
        raise SystemExit('this one-pager is already co-branded; start from the Trillium original')
    hits = blocked_hits(h)
    if hits:
        raise SystemExit(f'not built: the source one-pager has internal wording {hits}; use the owner one-pager')
    h = h.replace('fonts.googleapis.com/css2?family=Archivo', f'fonts.googleapis.com/css2?family={pr["font"]["google"]}&family=Archivo', 1)
    css = f'''
.pres{{background:{c['dark']};color:{c['text']};display:flex;justify-content:space-between;align-items:center;padding:.09in .55in;border-bottom:2px solid {c['accent']};font-family:'{F}','Archivo',sans-serif}}
.pres .k{{font-size:6.6pt;letter-spacing:.16em;text-transform:uppercase;color:{c['muted']};font-weight:600}}
.pres .n{{font-size:11pt;font-weight:700;letter-spacing:-.01em;margin-top:1px}}.pres .n i{{color:{c['accent']};font-style:normal}}
.pres .r{{text-align:right;font-size:7.6pt;color:{c['soft']};line-height:1.35}}.pres .r b{{color:{c['text']};font-weight:600}}
.pres img{{height:.28in;width:auto}}
.cobar{{display:grid;grid-template-columns:1fr 1fr;margin-top:auto}}.cobar .end{{margin-top:0;padding:.09in .4in .1in}}
.cobar .end b{{font-size:11pt}}.cobar .end .s{{font-size:8pt}}.cobar .end .u{{margin-top:.04in}}
.cobar .mia{{background:{c['dark']};color:{c['text']};padding:.09in .4in .1in;font-family:'{F}','Archivo',sans-serif;border-left:2px solid {c['accent']}}}
.mia .k{{font-size:6.6pt;letter-spacing:.16em;text-transform:uppercase;color:{c['accent']};font-weight:600}}
.mia .n{{font-size:12pt;font-weight:700;margin-top:3px;letter-spacing:-.01em}}
.mia .b{{font-size:7.4pt;letter-spacing:.12em;text-transform:uppercase;color:{c['muted']};font-weight:600;margin-top:1px}}
.mia .c{{font-size:8pt;color:{c['soft']};margin-top:.04in;line-height:1.4}}
.fn2{{font-size:6.1pt;color:#55645E;margin-top:.05in}}
.top{{padding:.2in .55in .16in}}.figs{{margin-top:.1in}}.fn{{font-size:6.3pt}}td{{padding:3px 8px}}
.big{{margin:.06in 0 .04in;font-size:12pt}}.q{{margin:.06in 0 .05in!important;font-size:13pt}}.set{{margin-top:.06in}}h2{{margin-top:.08in}}
'''
    h = h.replace('</style>', css + '</style>', 1)
    mark = f'<i>{esc(pr["brand_mark"])}</i>' if pr.get('brand_mark') else ''
    left = (f'<img src="/_blob/{esc(pr["logo_blob"])}" alt="{esc(pr["brand"])}">' if pr.get('logo_blob') else
            f'<div><div class="k">Presented by</div><div class="n">{esc(pr["name"])} <span style="font-weight:400;color:{c["muted"]}">·</span> {esc(pr["brand"])}{mark}</div></div>')
    pres = (f'<div class="pres">{left}<div class="r"><b>{esc(brokerage_line(pr))}</b><br>'
            f'{" · ".join(esc(x) for x in contacts(pr)[:2])}</div></div>')
    h = h.replace('<body><div class="page">', '<body><div class="page">' + pres, 1)
    h = re.sub(r'height:1\.\d+in;object-fit', 'height:.62in;object-fit:cover;object-position:center 60%', h, count=1)
    h = re.sub(r'<p class="cta">.*?</p>', '', h, count=1)
    i, j = h.index('<div class="end foot">'), h.rindex('</div></body>')
    sub = ' · '.join(x for x in [pr['brand'], pr.get('brokerage')] if x)
    mia = (f'<div class="mia"><div class="k">{esc(pr["presenter_label"])}</div><div class="n">{esc(pr["name"])}</div>'
           f'<div class="b">{esc(sub)}</div><div class="c">{"<br>".join(esc(x) for x in contacts(pr))}</div></div>')
    h = h[:i] + '<div class="cobar">' + h[i:j] + mia + '</div>' + h[j:]
    first = pr['name'].split()[0]
    disc = DISCLAIMER.format(name=pr['name'], brand=pr['brand'], brokerage_line=brokerage_line(pr), first=first)
    reg = f' {esc(pr["registration"])}' if pr.get('registration') else ''
    h = re.sub(r'(<p class="fn">.*?</p>)', lambda m: m.group(1) + f'<p class="fn2">{esc(disc)}{reg}</p>', h, count=1, flags=re.S)
    open(out, 'w').write(h)
    print(json.dumps({'out': out, 'warnings': check(pr)}, indent=1))


if __name__ == '__main__':
    a = sys.argv[1:]
    if not a or a[0] not in ('deck', 'onepager', 'check'):
        raise SystemExit(__doc__)
    if a[0] == 'check':
        w = check(load(a[1]))
        print('\n'.join(w) if w else 'profile OK')
    elif a[0] == 'deck':
        drop = DEFAULT_DROP
        if '--drop' in a:
            drop = a[a.index('--drop') + 1].split(',')
        rw = json.load(open(a[a.index('--rewrites') + 1])) if '--rewrites' in a else None
        deck(a[1], a[2], a[3], drop, rw)
    else:
        onepager(a[1], a[2], a[3])
