// 지옥철 서바이버 랭킹 서버 (Google Apps Script) v1
// 시트 이름: records  (첫 줄 머리글은 자동 생성)
const COLS=['time','nick','ver','diff','mode','daily','char','win','clock','kills','lv','dmg','endlessMin'];
function sheet_(){const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName('records');
  if(!sh){sh=ss.insertSheet('records');sh.appendRow(COLS);}return sh;}
function out_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){
  try{const d=JSON.parse(e.postData.contents);
    const nick=String(d.nick||'').slice(0,12);if(!nick)return out_({ok:false,err:'nick'});
    const clock=Number(d.clock)||0;if(d.win&&clock<1800)return out_({ok:false,err:'suspicious'}); // 30분 미만 클리어는 거름
    const lock=LockService.getScriptLock();lock.waitLock(5000);
    sheet_().appendRow([new Date(),nick,String(d.ver||''),String(d.diff||''),String(d.mode||''),String(d.daily||''),String(d.char||''),d.win?1:0,clock,Number(d.kills)||0,Number(d.lv)||0,Number(d.dmg)||0,Number(d.endlessMin)||0]);
    lock.releaseLock();return out_({ok:true});
  }catch(err){return out_({ok:false,err:String(err)});}
}
function doGet(e){
  const cat=(e.parameter.cat||'daily'),day=e.parameter.day||'';
  const v=sheet_().getDataRange().getValues();v.shift();
  let rows=v.map(r=>{const o={};COLS.forEach((c,i)=>o[c]=r[i]);return o;});
  let key,asc=false;
  if(cat==='daily'){rows=rows.filter(r=>r.mode==='daily'&&String(r.daily)===day);rows.sort((a,b)=>(b.win-a.win)||(b.win?a.clock-b.clock:b.clock-a.clock)||(b.kills-a.kills));}
  else if(cat==='normal'||cat==='hell'||cat==='easy'){rows=rows.filter(r=>r.diff===cat&&r.win==1&&r.mode==='normal');key='clock';asc=true;}
  else if(cat==='kills'){key='kills';}
  else if(cat==='lv'){key='lv';}
  else if(cat==='endless'){rows=rows.filter(r=>r.mode==='endless');key='endlessMin';}
  if(key)rows.sort((a,b)=>asc?a[key]-b[key]:b[key]-a[key]);
  // 닉네임당 최고 기록 1개만
  const seen={},top=[];for(const r of rows){if(seen[r.nick])continue;seen[r.nick]=1;top.push(r);if(top.length>=20)break;}
  return out_({ok:true,cat,rows:top});
}
