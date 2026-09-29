# Excel(類型A/B/C) -> spec JSON 抽出スクリプト
import re, glob, json, sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KEYS = {'類型A':'a','類型B':'b','類型C':'c'}
REF = re.compile(r"INDEX\('([^']+)'!\$E:\$E,(\d+)\),\"　\",\" \"\)\)(=|<>)\"([^\"]*)\",\"(対象外|先に選択)\"")
def load(path):
    wb = openpyxl.load_workbook(path)
    sheets = []
    for ws in wb:
        rows = {}
        for r in ws.iter_rows():
            for c in r: 
                if c.value is not None: rows.setdefault(c.row, {})[c.column_letter] = c.value
        dvs = {}
        for dv in ws.data_validations.dataValidation:
            for rng in str(dv.sqref).split():
                dvs[rng] = dv.formula1.strip('"') if dv.formula1 else ''
        sheets.append((ws.title, rows, dvs))
    return sheets
def parse(path):
    sheets = load(path)
    idmap = {}  # (sheet,row)->id
    for t, rows, _ in sheets:
        for n, r in rows.items():
            if isinstance(r.get('A'), str) and re.match(r'^[A-Z]\d+-\d+$|^\d\d-\d\d$', r['A']): idmap[(t, n)] = r['A']
    out = {'sheets': [], 'raw': {}}
    for t, rows, dvs in sheets:
        qs = []
        for n in sorted(rows):
            r = rows[n]
            if (t, n) not in idmap: continue
            f = r.get('F', '')
            deps = []
            for m in REF.finditer(str(f)):
                deps.append({'ref': idmap.get((m.group(1), int(m.group(2))), '?%s!%s' % (m.group(1), m.group(2))), 'op': m.group(3), 'val': m.group(4), 'kind': m.group(5)})
            if str(f).count('対象外') != sum(1 for d in deps if d['kind']=='対象外') : print('UNPARSED', r['A'], f)
            opts = dvs.get('E%d' % n)
            if not opts:
                for k, v in dvs.items():
                    if k == 'E%d' % n: opts = v
            qs.append({'id': r['A'], 'label': r.get('B', ''), 'prompt': r.get('C', ''), 'example': r.get('D', ''), 'options': opts, 'deps': deps, 'row': n})
        entry = {'title': t, 'head': rows.get(1, {}).get('A'), 'note': rows.get(3, {}).get('A'), 'qs': qs}
        if t == '提出書類チェック':
            entry['docs'] = [{'id': 'D%d' % r['A'], 'no': r['A'], 'label': r['B']} for n, r in sorted(rows.items()) if n >= 5 and isinstance(r.get('A'), int)]
            entry['docStates'] = [x.strip() for x in dvs.get('C5', '').split(',')] if dvs.get('C5') else []
        out['sheets'].append(entry)
    return out, sheets
if __name__ == '__main__':
    for f in sorted(glob.glob(os.path.join(ROOT, '*.xlsx'))):
        k = [v for kk, v in KEYS.items() if kk in f][0]
        spec, _ = parse(f)
        json.dump(spec, open(os.path.join(ROOT, 'tools', 'spec_%s.json' % k), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        print(k, [(s['title'], len(s['qs'])) for s in spec['sheets']])
