"""Read the source workbook without modifying it. Run with Python + openpyxl + Pillow."""
import json, sys
from pathlib import Path
from datetime import datetime, date
import openpyxl
from PIL import Image
from io import BytesIO

root = Path(__file__).resolve().parents[1]
source = Path(sys.argv[1]) if len(sys.argv)>1 else root.parent/'Japan_Product_Knowledge_Catalog.xlsx'
book = openpyxl.load_workbook(source, data_only=True)
sheet = book['Product_Knowledge']
out = root/'public'/'images'
out.mkdir(parents=True, exist_ok=True)
images = {}
for img in sheet._images:
    row = img.anchor._from.row + 1
    name = f'product-{row}.webp'
    picture = Image.open(BytesIO(img._data())).convert('RGB')
    picture.thumbnail((700,700))
    picture.save(out/name, 'WEBP', quality=85)
    images.setdefault(row, []).append('/images/'+name)
fields = ['number','imageNote','barcode','itemCode','name','department','category','knowledge','uom','price','expiry','action']
products=[]
for row, values in enumerate(sheet.iter_rows(min_row=4,values_only=True),4):
    if not values[4]: continue
    p=dict(zip(fields,values))
    for key,value in p.items():
        if isinstance(value,(datetime,date)): p[key]=value.strftime('%Y-%m-%d')
        elif value is None: p[key]=''
    p['id']=str(p['number'])
    p['barcode']=str(p['barcode']).strip()
    p['itemCode']=str(p['itemCode']).strip()
    p['images']=images.get(row,[])
    products.append(p)
notes=[{'label':r[0] or '', 'text':r[1] or ''} for r in book['Notes'].iter_rows(values_only=True) if r[0]]
data={'products':products,'notes':notes,'source':source.name,'priceDate':'2026-10-30'}
(root/'public'/'catalog.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'products':len(products),'images':len(images),'departments':sorted(set(p['department'] for p in products))}))
