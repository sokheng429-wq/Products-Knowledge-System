export const normalizeBarcode = value => String(value ?? '').replace(/[\s-]/g, '');
export function barcodeMatches(products, code) {
 const value=normalizeBarcode(code); if(!value)return [];
 return products.filter(p=>normalizeBarcode(p.barcode)===value||(value.length===12&&normalizeBarcode(p.barcode)==='0'+value)||(value.length===13&&value[0]==='0'&&normalizeBarcode(p.barcode)===value.slice(1)));
}
export function expiryStatus(value, now=new Date()) {
 if(!/^\d{4}-\d{2}-\d{2}$/.test(value))return null;
 const today=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Bangkok',year:'numeric',month:'2-digit',day:'2-digit'}).format(now);
 const days=Math.round((Date.parse(value+'T00:00:00Z')-Date.parse(today+'T00:00:00Z'))/86400000);
 return days<0?'Expired':days<=45?'Expires soon':null;
}
export function filterProducts(products,query,department,category) {
 const term=query.trim().toLocaleLowerCase();
 return products.filter(p=>(department==='All products'||p.department===department)&&(!category||p.category===category)&&(!term||[p.name,p.barcode,p.itemCode,p.category,p.knowledge].some(v=>String(v).toLocaleLowerCase().includes(term))));
}
