import ExcelJS from 'exceljs';
import {normalizeBarcode} from './catalog.js';

const columns = [
 ['number','No.',8],['imageNote','Image',24],['barcode','Barcode',22],['itemCode','Item Code',16],
 ['name','Product Name',44],['department','Department',25],['category','Category',25],
 ['knowledge','Product Knowledge',70],['uom','UOM',12],['price','Price (USD)',16],
 ['expiry','Expiry',16],['action','Action Needed',40],
];
function text(value) {
 if(value==null)return '';
 if(value instanceof Date)return value.toISOString().slice(0,10);
 if(typeof value==='object') {
  if('formula' in value || 'sharedFormula' in value)return text(value.result);
  if(value.richText)return value.richText.map(t=>t.text).join('');
  return value.text||'';
 }
 return String(value).trim();
}
function key(value){return text(value).split('\n').at(-1).toLowerCase().replace(/[^a-z0-9]/g,'');}
function imageData(media){
 if(media.base64)return media.base64.startsWith('data:')?media.base64:`data:image/${media.extension};base64,${media.base64}`;
 const bytes=new Uint8Array(media.buffer);let str='';
 for(let i=0;i<bytes.length;i+=8192)str+=String.fromCharCode(...bytes.subarray(i,i+8192));
 return `data:image/${media.extension};base64,${btoa(str)}`;
}
export async function parseWorkbook(buffer,source='Imported catalog.xlsx') {
 const book=new ExcelJS.Workbook();
 try{await book.xlsx.load(buffer);}catch{throw new Error('This file could not be read. Choose a valid .xlsx workbook (not .xls, CSV, or a password-protected file).');}
 let sheet,header,mapping;
 for(const candidate of book.worksheets){
  for(let r=1;r<=Math.min(candidate.rowCount,20);r++){
   const map={};candidate.getRow(r).eachCell((cell,c)=>{const match=columns.find(([,label])=>key(label)===key(cell.value));if(match)map[match[0]]=c;});
   if(map.barcode&&map.name){sheet=candidate;header=r;mapping=map;break;}
  }if(sheet)break;
 }
 if(!sheet)throw new Error('Missing required columns: Barcode and Product Name. Use the template or your original Product_Knowledge workbook.');
 if(sheet.rowCount-header>5000)throw new Error('Import up to 5,000 product rows at a time.');
 const pictures=new Map();
 for(const img of sheet.getImages()){
  const media=book.getImage(img.imageId);
  if(!media||!['png','jpeg','jpg','gif'].includes(media.extension))continue;
  const row=Math.floor(img.range.tl.nativeRow??img.range.tl.row)+1;
  if(!pictures.has(row))pictures.set(row,[]);
  pictures.get(row).push(imageData(media));
 }
 const products=[],errors=[],seen=new Set();let numericBarcodes=0;
 for(let r=header+1;r<=sheet.rowCount;r++){
  const row=sheet.getRow(r);if(!row.hasValues)continue;
  const p={};for(const [field]of columns)p[field]=mapping[field]?text(row.getCell(mapping[field]).value):'';
  if(!p.name&&!p.barcode)continue;
  p.barcode=normalizeBarcode(p.barcode);
  if(!p.name){errors.push(`Row ${r}: product name is required.`);continue;}
  if(!/^\d{8,14}$/.test(p.barcode)){errors.push(`Row ${r}: barcode must contain 8–14 digits. Format barcode cells as Text.`);continue;}
  if(seen.has(p.barcode)){errors.push(`Row ${r}: duplicate barcode ${p.barcode}.`);continue;}seen.add(p.barcode);
  if(typeof row.getCell(mapping.barcode).value==='number')numericBarcodes++;
  if(p.price!=='') {const price=Number(p.price);if(!Number.isFinite(price)||price<0){errors.push(`Row ${r}: price must be a non-negative number or blank.`);continue;}p.price=price;}
  if(p.expiry&&(!/^\d{4}-\d{2}-\d{2}$/.test(p.expiry)||!Number.isFinite(Date.parse(p.expiry))||new Date(p.expiry).toISOString().slice(0,10)!==p.expiry)){errors.push(`Row ${r}: expiry must be an Excel date or YYYY-MM-DD.`);continue;}
  p.id=p.barcode;p.department=p.department||'UNCATEGORIZED';p.category=p.category||'Uncategorized';p.images=pictures.get(r)||[];products.push(p);
 }
 if(!products.length&&!errors.length)throw new Error('The workbook contains no product rows. Add products below the header and try again.');
 const notes=[];book.getWorksheet('Notes')?.eachRow(row=>{const label=text(row.getCell(1).value),value=text(row.getCell(2).value);if(label)notes.push({label,text:value});});
 const metadata=book.getWorksheet('Catalog_Info');
 const priceDate=metadata?text(metadata.getCell('B1').value):'';
 return {products,notes,source,priceDate,errors,numericBarcodes,fields:Object.keys(mapping),imageCount:products.filter(p=>p.images.length).length};
}
export function combineCatalog(current,incoming,mode='merge') {
 const existing=new Map(current.products.map(p=>[p.barcode,p]));
 const incomingProducts=incoming.products.map(p=>{
  const previous=existing.get(p.barcode);
  const next=mode==='merge'&&previous&&incoming.fields?{...previous,...Object.fromEntries(incoming.fields.map(field=>[field,p[field]]))}:{...p};
  return {...next,id:p.barcode,images:p.images.length?p.images:(previous?.images||[])};
 });
 const result=mode==='replace'?incomingProducts:[...new Map([...current.products,...incomingProducts].map(p=>[p.barcode,p])).values()];
 return {products:result,notes:incoming.notes.length?incoming.notes:current.notes,source:incoming.source,priceDate:incoming.priceDate||'',updatedAt:new Date().toISOString()};
}
export async function createWorkbook(data,{products=data.products,includeImages=false,resolveImage,onProgress=()=>{}}={}) {
 const book=new ExcelJS.Workbook();book.creator='B’Groceries';
 const sheet=book.addWorksheet('Product_Knowledge',{views:[{state:'frozen',ySplit:1}]});
 sheet.columns=columns.map(([key,header,width])=>({key,header,width}));
 const head=sheet.getRow(1);head.height=30;head.eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF153F34'}};c.font={bold:true,color:{argb:'FFFFFFFF'},size:11};c.alignment={vertical:'middle'};});
 sheet.getColumn('barcode').numFmt='@';sheet.getColumn('itemCode').numFmt='@';sheet.getColumn('price').numFmt='"$"#,##0.00';
 let missingImages=0;
 for(let i=0;i<products.length;i++){
  const p=products[i];const row=sheet.addRow({...p,number:p.number||i+1,imageNote:p.imageNote||'',price:p.price===''?null:p.price});row.height=includeImages?100:64;
  row.eachCell(c=>{c.font={name:'Calibri',size:11,color:{argb:'FF233B34'}};c.alignment={vertical:'top',wrapText:true};if(i%2===1)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF1F5EE'}};});
  if(includeImages&&p.images.length){try{const base64=resolveImage?await resolveImage(p.images[0]):p.images[0];if(!/^data:image\/(png|jpeg|gif);base64,/.test(base64))throw new Error('Unsupported image');const extension=base64.match(/^data:image\/(\w+)/)[1];const id=book.addImage({base64,extension});sheet.addImage(id,{tl:{col:1.1,row:row.number-0.95},ext:{width:115,height:115},editAs:'oneCell'});}catch{missingImages++;}}
  if(i%10===0){onProgress(Math.round((i+1)/Math.max(products.length,1)*100));await new Promise(resolve=>setTimeout(resolve,0));}
 }
 sheet.autoFilter={from:'A1',to:`L${Math.max(1,sheet.rowCount)}`};
 const notes=book.addWorksheet('Notes');notes.columns=[{width:45},{width:100}];for(const n of data.notes||[])notes.addRow([n.label,n.text]);notes.eachRow(r=>r.eachCell(c=>c.alignment={wrapText:true,vertical:'top'}));
 const info=book.addWorksheet('Catalog_Info');info.columns=[{width:24},{width:70}];info.addRow(['Price date',data.priceDate||'']);info.addRow(['Source',data.source||'']);info.addRow(['Instructions','Keep barcodes and item codes as Text. Required: Barcode and Product Name. Expiry: YYYY-MM-DD or Excel date.']);
 return {buffer:await book.xlsx.writeBuffer(),missingImages};
}
export async function browserImage(source){
 if(!source.startsWith('/images/')&&!/^data:image\/(png|jpeg|jpg|gif|webp);base64,/.test(source))throw new Error('Unsupported image source');
 const image=new Image();image.src=source;await image.decode();
 const canvas=document.createElement('canvas');canvas.width=320;canvas.height=320;const scale=320/Math.max(image.width,image.height);const w=image.width*scale,h=image.height*scale;canvas.getContext('2d').drawImage(image,(320-w)/2,(320-h)/2,w,h);return canvas.toDataURL('image/png');
}
