import fs from 'node:fs';
const original=fs.readFileSync(new URL('source/app.js',import.meta.url),'utf8');
function change(from,to){if(!app.includes(from))throw new Error(`Cannot find: ${from.slice(0,50)}`);app=app.replace(from,to)}
let app=original;
change('const data=window.SALES_DATA||[],money=', 'let data=[];const money=');
change('const products=[...itemTotals.keys()]', 'let products=[...itemTotals.keys()]');
change('new Set(saved.filter(x=>products.includes(x)))','new Set()');
change('let saved=[];try{saved=JSON.parse(localStorage.getItem("chenghuang-selected-products")||"[]")}catch(_){ }','let saved=[];');
app=app.replaceAll('localStorage.setItem("chenghuang-selected-products",JSON.stringify([...state.selected]))','void 0').replaceAll('localStorage.removeItem("chenghuang-selected-products")','void 0');
change('state.compareMonth=String(data.at(-1)?.gm||1);renderControls(true);render();', '$("excel-upload").addEventListener("change",async event=>{const file=event.target.files?.[0];if(!file)return;const status=$("upload-status");status.textContent="正在讀取與檢查資料…";try{const result=await window.readSalesWorkbook(file);data=result.records;itemTotals.clear();data.forEach(r=>itemTotals.set(r.item,(itemTotals.get(r.item)||0)+r.total));products=[...itemTotals.keys()].sort((a,b)=>itemTotals.get(b)-itemTotals.get(a));state.selected=new Set(saved.filter(x=>products.includes(x)));state.heatmapProduct=[...state.selected][0]||"";state.calendar="gregorian";state.month="all";state.compareMonth=String(data.at(-1)?.gm||1);renderControls(true);render();status.textContent=`已在本機載入 ${file.name}：${data.length} 筆、${products.length} 項商品；總價修正 ${result.corrected} 筆、略過 ${result.bad} 筆。未上傳到網站。`}catch(error){status.textContent=`讀取失敗：${error.message}`}});');
fs.writeFileSync(new URL('dist/app.js',import.meta.url),app);
