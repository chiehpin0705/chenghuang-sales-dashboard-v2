window.readSalesWorkbook=async function(file){
  if(!/\.xlsx$/i.test(file.name))throw new Error('請選擇 .xlsx 檔案');
  if(file.size>20*1024*1024)throw new Error('檔案超過 20 MB');
  const buffer=await file.arrayBuffer(), view=new DataView(buffer), bytes=new Uint8Array(buffer);
  let end=-1;for(let p=bytes.length-22;p>=Math.max(0,bytes.length-65557);p--)if(view.getUint32(p,true)===0x06054b50){end=p;break}
  if(end<0)throw new Error('Excel 檔案格式無法辨識');
  const entries=new Map(),count=view.getUint16(end+10,true);let pos=view.getUint32(end+16,true);
  for(let i=0;i<count;i++){
    if(view.getUint32(pos,true)!==0x02014b50)throw new Error('Excel 壓縮資料不完整');
    const method=view.getUint16(pos+10,true),size=view.getUint32(pos+20,true),nameLen=view.getUint16(pos+28,true),extra=view.getUint16(pos+30,true),comment=view.getUint16(pos+32,true),offset=view.getUint32(pos+42,true),name=new TextDecoder().decode(bytes.slice(pos+46,pos+46+nameLen));
    entries.set(name,{method,size,offset});pos+=46+nameLen+extra+comment;
  }
  async function part(name){const e=entries.get(name);if(!e)return null;const p=e.offset;if(view.getUint32(p,true)!==0x04034b50)throw new Error('Excel 檔案損壞');const start=p+30+view.getUint16(p+26,true)+view.getUint16(p+28,true),compressed=bytes.slice(start,start+e.size);let data;
    if(e.method===0)data=compressed;
    else if(e.method===8){if(!globalThis.DecompressionStream)throw new Error('請使用最新版 Chrome、Edge 或 Safari 開啟');data=new Uint8Array(await new Response(new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())}
    else throw new Error('不支援的 Excel 壓縮格式');
    return new DOMParser().parseFromString(new TextDecoder().decode(data),'application/xml');
  }
  const children=(node,name)=>Array.from(node?.children||[]).filter(x=>x.localName===name),first=(node,name)=>children(node,name)[0],value=node=>node?.textContent?.trim()||'';
  const stringsDoc=await part('xl/sharedStrings.xml'),strings=children(stringsDoc?.documentElement,'si').map(si=>Array.from(si.getElementsByTagName('*')).filter(x=>x.localName==='t').map(value).join(''));
  const sheet=await part('xl/worksheets/sheet1.xml');if(!sheet)throw new Error('找不到第一個工作表');
  const sheetData=first(sheet.documentElement,'sheetData'),rows=children(sheetData,'row').map(row=>{
    const cells={};for(const c of children(row,'c')){const col=(c.getAttribute('r')||'').match(/^[A-Z]+/)?.[0];if(!col)continue;let v=value(first(c,'v'));if(c.getAttribute('t')==='s')v=strings[Number(v)]||'';else if(c.getAttribute('t')==='inlineStr')v=value(first(c,'is'));cells[col]=v}return cells;
  });
  if(rows.length<2)throw new Error('工作表沒有銷售資料');
  const header=rows[0],columns=Object.fromEntries(Object.entries(header).map(([k,v])=>[v.replace(/\s/g,''),k]));
  for(const h of ['西元年月日','品項','數量','單位','單價','總價'])if(!columns[h])throw new Error(`缺少「${h}」欄`);
  const lunarFormat=new Intl.DateTimeFormat('zh-TW-u-ca-chinese',{year:'numeric',month:'long',day:'numeric'}),lunarDigits={'正':1,'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,'冬':11,'十一':11,'臘':12,'十二':12};
  let currentDate='',bad=0,corrected=0;const records=[];
  function parseDate(v){if(!v)return currentDate;let date;if(/^\d+(\.\d+)?$/.test(v))date=new Date(Date.UTC(1899,11,30)+Math.round(Number(v))*86400000);else{const m=v.match(/^(\d{4})[-\/]([01]?\d)[-\/]([0-3]?\d)/);if(!m)return '';date=new Date(Date.UTC(+m[1],+m[2]-1,+m[3]))}return Number.isNaN(date.getTime())?'':date.toISOString().slice(0,10)}
  rows.slice(1).forEach(row=>{if(!Object.values(row).some(Boolean))return;const d=parseDate(row[columns['西元年月日']]);if(d)currentDate=d;const item=row[columns['品項']]?.trim(),qty=Number(row[columns['數量']]),price=Number(row[columns['單價']]),unit=row[columns['單位']]?.trim()||'',entered=Number(row[columns['總價']]);
    if(!d||!item||!Number.isFinite(qty)||!Number.isFinite(price)){bad++;return}
    const total=qty*price;if(Number.isFinite(entered)&&Math.abs(entered-total)>0.01)corrected++;
    const date=new Date(`${d}T12:00:00Z`),parts=lunarFormat.formatToParts(date),ly=Number(parts.find(p=>p.type==='relatedYear')?.value),monthText=parts.find(p=>p.type==='month')?.value||'',leap=monthText.includes('閏'),lm=lunarDigits[monthText.replace(/[閏月]/g,'')];
    if(!ly||!lm){bad++;return}records.push({d,gy:date.getUTCFullYear(),gm:date.getUTCMonth()+1,ly,lm,leap,item,qty,unit,price,total});
  });
  if(!records.length)throw new Error('找不到可分析的銷售紀錄');
  records.sort((a,b)=>a.d.localeCompare(b.d));return{records,corrected,bad};
};
