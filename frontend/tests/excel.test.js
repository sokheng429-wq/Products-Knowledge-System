import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import ExcelJS from 'exceljs';
import {parseWorkbook,createWorkbook,combineCatalog} from '../src/excel.js';
const original=JSON.parse(readFileSync(new URL('../public/catalog.json',import.meta.url)));
const picture='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jN1sAAAAASUVORK5CYII=';
test('original bilingual workbook imports all rows and embedded pictures',async()=>{
 const result=await parseWorkbook(readFileSync(new URL('../../Japan_Product_Knowledge_Catalog.xlsx',import.meta.url)));
 assert.deepEqual(result.errors,[]);assert.equal(result.products.length,797);assert.equal(result.imageCount,791);
 for(let i=0;i<797;i++){for(const field of ['barcode','name','knowledge','price','itemCode','expiry','uom','action'])assert.equal(result.products[i][field],original.products[i][field],`${i}: ${field}`);}
});
test('export reimports identifiers, Khmer, literal text, prices, dates, notes and photos',async()=>{
 const p={...original.products[0],barcode:'0012345678901',itemCode:'0007',name:'=literal text, not a formula',images:[picture]};
 const {buffer,missingImages}=await createWorkbook({...original,products:[p]},{includeImages:true});
 const result=await parseWorkbook(buffer);assert.equal(missingImages,0);assert.deepEqual(result.errors,[]);
 for(const field of ['barcode','itemCode','name','knowledge','price','expiry'])assert.equal(result.products[0][field],p[field]);
 assert.equal(result.imageCount,1);assert.equal(result.priceDate,original.priceDate);assert.deepEqual(result.notes,original.notes);
 const wb=new ExcelJS.Workbook();await wb.xlsx.load(buffer);assert.equal(wb.worksheets[0].getCell('E2').type,ExcelJS.ValueType.String);
});
test('rejects malformed, missing headers, empty workbooks and invalid rows',async()=>{
 await assert.rejects(()=>parseWorkbook(new Uint8Array([1,2,3])),/could not be read/);
 const book=new ExcelJS.Workbook();const sheet=book.addWorksheet('Products');sheet.addRow(['Product Name','Barcode','Price (USD)','Expiry']);
 await assert.rejects(()=>parseWorkbookFrom(book),/no product rows/);
 sheet.addRow(['Coffee','0012345678901',2.5,'2026-11-30']);sheet.addRow(['Duplicate','0012345678901',3,'']);sheet.addRow(['Bad price','0012345678902',-1,'']);sheet.addRow(['Bad date','0012345678903',1,'2026-02-31']);sheet.addRow(['','0012345678904',1,'']);
 const result=await parseWorkbookFrom(book);assert.equal(result.products.length,1);assert.equal(result.errors.length,4);
 const missing=new ExcelJS.Workbook();missing.addWorksheet('Data').addRow(['foo']);await assert.rejects(()=>parseWorkbookFrom(missing),/Missing required columns/);
});
test('merge updates by barcode and preserves images; replace removes other rows',()=>{
 const current={...original,products:original.products.slice(1,3)};const update={...current.products[0],name:'Updated',images:[]};
 const incoming={products:[update],notes:[],source:'update.xlsx'};
 const merged=combineCatalog(current,incoming);assert.equal(merged.products.length,2);assert.equal(merged.products[0].name,'Updated');assert.deepEqual(merged.products[0].images,current.products[0].images);assert.equal(merged.priceDate,'');
 assert.equal(combineCatalog(current,incoming,'replace').products.length,1);assert.equal(current.products[0].name,original.products[1].name);
});
test('filtered export includes only selected rows and template keeps text-formatted barcode column',async()=>{
 const {buffer}=await createWorkbook(original,{products:original.products.slice(0,2)});const result=await parseWorkbook(buffer);assert.equal(result.products.length,2);
 const template=await createWorkbook(original,{products:[]});const book=new ExcelJS.Workbook();await book.xlsx.load(template.buffer);assert.equal(book.worksheets[0].getColumn(3).numFmt,'@');assert.equal(book.worksheets[0].rowCount,1);
});
test('partial imports preserve omitted columns while explicit blanks clear fields',()=>{
 const current={...original,products:[original.products[0]]};
 const incoming={products:[{...original.products[0],name:'Changed',knowledge:'',price:''}],fields:['barcode','name','price'],notes:[],source:'partial.xlsx'};
 const merged=combineCatalog(current,incoming);assert.equal(merged.products[0].name,'Changed');assert.equal(merged.products[0].price,'');assert.equal(merged.products[0].knowledge,original.products[0].knowledge);
});
async function parseWorkbookFrom(book){return parseWorkbook(await book.xlsx.writeBuffer());}
