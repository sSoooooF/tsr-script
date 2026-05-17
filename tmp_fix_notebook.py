import json
from pathlib import Path
path = Path('tsr/tsr_audit.ipynb')
nb = json.loads(path.read_text(encoding='utf-8'))
for cell in nb.get('cells', []):
    if cell.get('id') == '#VSC-ef951f35':
        print('CELL TYPE', cell.get('cell_type'))
        source = ''.join(cell.get('source', []))
        print('SOURCE START')
        print(source)
        print('SOURCE END')
        if "current_start = 2    print('Excel export failed:'" in source:
            print('FOUND MALFORMED LINE')
            source = source.replace("current_start = 2    print('Excel export failed:', exc)", "current_start = 2")
            cell['source'] = source.splitlines(keepends=True)
            path.write_text(json.dumps(nb, ensure_ascii=False, indent=1), encoding='utf-8')
            print('FIXED MALFORMED LINE')
        break
else:
    print('CELL NOT FOUND')
