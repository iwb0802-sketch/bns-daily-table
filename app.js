let rawRows = [];
let mappedRows = [];
const COLS = {
  no:['No.','No','번호'],
  date:['행사날짜'],
  time:['시간'],
  place:['장소/층수','장소','장소/충수'],
  order:['발주처'],
  arrangement:['연주편성'],
  payTotal:['연주자페이합계'],
};

const $ = (id)=>document.getElementById(id);
const fileInput=$('fileInput'), dropZone=$('dropZone'), statusEl=$('fileStatus');
$('fileBtn').onclick=()=>fileInput.click();
fileInput.onchange=(e)=>handleFile(e.target.files[0]);
['dragenter','dragover'].forEach(evt=>dropZone.addEventListener(evt,e=>{e.preventDefault();dropZone.classList.add('drag')}));
['dragleave','drop'].forEach(evt=>dropZone.addEventListener(evt,e=>{e.preventDefault();dropZone.classList.remove('drag')}));
dropZone.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));
$('makeBtn').onclick=render;
$('downloadBtn').onclick=downloadExcel;
$('dateSelect').onchange=render;
$('searchInput').oninput=()=>render();
$('hideCancel').onchange=render;

function normalize(s){return String(s??'').replace(/\s+/g,' ').trim();}
function money(v){const n=Number(String(v??'').replace(/[^0-9.-]/g,''));return isNaN(n)?0:n;}
function findCol(row, names){const keys=Object.keys(row); for(const n of names){const k=keys.find(x=>normalize(x)===n); if(k) return k;} return null;}
function parseTimeValue(t){
  // BNS 시간 문자열 예: "11시", "12시30분", "1시", "3시40분", "10시30분~6시".
  // 웨딩 진행표는 보통 오전 10~12시 이후 오후 1~7시 순서라서
  // 1~7시는 오후 시간(13~19시)으로 보정해 정렬한다.
  t=normalize(t).replace(/\(취소\)/g,'').replace(/취소/g,'');
  const m=t.match(/(\d{1,2})\s*시\s*(?:(\d{1,2})\s*분?)?/);
  if(!m) return 9999;
  let h=parseInt(m[1],10);
  const min=m[2]?parseInt(m[2],10):0;
  if(h>=1 && h<=7) h+=12;
  return h*60+min;
}
function dateOnly(v){return normalize(v).replace(/\(.+?\)/g,'');}
function stripRoleAndName(text){
  text=normalize(text).replace(/[,，]?\s*\d{1,3}(,\d{3})+\s*원?/g,'').replace(/\d+\s*원/g,'').trim();
  return text;
}
function performerKey(text){
  const clean=stripRoleAndName(text);
  if(!clean) return '';
  // 중복 표시는 이름 기준. '김영일 사회 60,000' -> '김영일'
  return clean.split(/[\s/]+/)[0];
}
function isCancel(row){
  return [row.time,row.place,row.order,row.arrangement].some(v=>normalize(v).includes('취소'));
}
function isYedoOnly(row){
  // 연주편성이 정확히 '예도'인 경우만 제외. '예도+사회' 같은 조합은 유지.
  return normalize(row.arrangement).replace(/\s+/g,'') === '예도';
}

async function handleFile(file){
  if(!file) return;
  try{
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array'});
    const ws=wb.Sheets[wb.SheetNames[0]];
    rawRows=XLSX.utils.sheet_to_json(ws,{defval:''});
    mappedRows=mapRows(rawRows);
    populateDates();
    statusEl.textContent=`업로드 완료: ${file.name} / ${mappedRows.length}행`;
    render();
  }catch(err){
    console.error(err);
    statusEl.textContent='파일을 읽지 못했습니다. BNS에서 받은 엑셀 파일인지 확인해주세요.';
    statusEl.style.color='#b00020';
  }
}

function mapRows(rows){
  return rows.map((r,i)=>{
    const get=(arr)=>{const k=findCol(r,arr); return k?r[k]:''};
    const row={
      no:get(COLS.no)||i+1,
      date:get(COLS.date),
      time:get(COLS.time),
      place:get(COLS.place),
      order:get(COLS.order),
      arrangement:get(COLS.arrangement),
      payTotal:get(COLS.payTotal),
      players:[]
    };
    for(let n=1;n<=10;n++){
      const k=findCol(r,[`악기구성${n}`]);
      row.players.push(k?normalize(r[k]):'');
    }
    return row;
  }).filter(r=>r.date||r.time||r.place||r.arrangement||r.players.some(Boolean));
}
function populateDates(){
  const sel=$('dateSelect'); const current=sel.value;
  const dates=[...new Set(mappedRows.map(r=>dateOnly(r.date)).filter(Boolean))].sort();
  sel.innerHTML='<option value="">전체 날짜</option>'+dates.map(d=>`<option value="${d}">${d}</option>`).join('');
  if(dates.includes(current)) sel.value=current;
}
function filteredRows(){
  const date=$('dateSelect').value;
  const q=normalize($('searchInput').value).toLowerCase();
  const hide=$('hideCancel').checked;
  let rows=mappedRows.filter(r=>!date||dateOnly(r.date)===date);
  if(hide) rows=rows.filter(r=>!isCancel(r));
  rows=rows.filter(r=>!isYedoOnly(r));
  if(q) rows=rows.filter(r=>[r.date,r.time,r.place,r.order,r.arrangement,...r.players].join(' ').toLowerCase().includes(q));
  return rows.sort((a,b)=>String(dateOnly(a.date)).localeCompare(String(dateOnly(b.date)),'ko') || parseTimeValue(a.time)-parseTimeValue(b.time) || String(a.place).localeCompare(String(b.place),'ko'));
}

function getActiveDates(rows){
  return [...new Set(rows.map(r=>dateOnly(r.date)).filter(Boolean))];
}
function render(){
  const rows=filteredRows();
  const multipleDates=getActiveDates(rows).length>1;
  let currentDate='';
  let seen=new Set();
  let dupCount=0;
  let displayNo=0;
  const maxPlayers=Math.min(10, Math.max(5,...rows.map(r=>r.players.reduce((m,p,i)=>p?i+1:m,0))));
  const headers=['No.','행사날짜','시간','장소/층수','연주편성',...Array.from({length:maxPlayers},(_,i)=>`악기구성${i+1}`)];
  const html=['<thead><tr>'+headers.map(h=>`<th>${h}</th>`).join('')+'</tr></thead><tbody>'];
  rows.forEach((r)=>{
    const d=dateOnly(r.date);
    if(multipleDates && currentDate && d!==currentDate){
      html.push(`<tr class="date-gap"><td colspan="${headers.length}"></td></tr>`);
      seen=new Set();
    }
    if(d!==currentDate){
      currentDate=d;
      if(multipleDates) seen=new Set();
    }
    displayNo++;
    const cancel=isCancel(r);
    html.push(`<tr class="${cancel?'cancel':''}">`);
    html.push(`<td class="num">${displayNo}</td><td class="date">${normalize(r.date)}</td><td class="time">${normalize(r.time)}</td><td class="place">${normalize(r.place)}</td><td class="arrangement">${normalize(r.arrangement)}</td>`);
    for(let i=0;i<maxPlayers;i++){
      const p=normalize(r.players[i]); const key=performerKey(p);
      let cls='player';
      if(key){ if(seen.has(key)){ cls = (cls ? cls + ' ' : '') + 'dup'; dupCount++; } else seen.add(key); }
      html.push(`<td class="${cls}">${p||''}</td>`);
    }
    html.push('</tr>');
  });
  html.push('</tbody>');
  $('resultTable').innerHTML=html.join('');
  $('tableInfo').textContent=`${rows.length}건 / 날짜 ${getActiveDates(rows).length}개 / 중복 표시 ${dupCount}칸`;
  $('summary').innerHTML=[
    ['행사 수',`${rows.length}건`],
    ['날짜 수',`${getActiveDates(rows).length}개`],
    ['총 연주자 입력칸',`${rows.reduce((s,r)=>s+r.players.filter(Boolean).length,0)}명`],
    ['중복 표시',`${dupCount}칸`]
  ].map(([l,v])=>`<div class="item"><div class="label">${l}</div><div class="value">${v}</div></div>`).join('');
}
function escapeHtml(value){
  return String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}
function downloadExcel(){
  const rows=filteredRows();
  if(!rows.length) return alert('다운로드할 데이터가 없습니다.');
  const multipleDates=getActiveDates(rows).length>1;
  let currentDate='';
  let seen=new Set();
  let displayNo=0;
  const maxPlayers=Math.min(10, Math.max(5,...rows.map(r=>r.players.reduce((m,p,i)=>p?i+1:m,0))));
  const headers=['No.','행사날짜','시간','장소/층수','연주편성',...Array.from({length:maxPlayers},(_,i)=>`악기구성${i+1}`)];
  const colWidths=[48,120,110,240,240,...Array.from({length:maxPlayers},()=>150)];
  let html=`<!doctype html><html><head><meta charset="UTF-8"><style>
    table{border-collapse:collapse;font-family:Arial,'맑은 고딕',sans-serif;font-size:12px;color:#111;}
    th,td{border:1px solid #999;padding:8px 7px;white-space:nowrap;vertical-align:middle;color:#111;text-decoration:none;height:32px;mso-height-source:userset;}
    tr{height:32px;mso-height-source:userset;}
    th{background:#eef1f4;font-weight:700;text-align:center;color:#111;}
    .center{text-align:center;color:#111;}
    .normal{color:#111;text-decoration:none;}
    .dup{background:#fff1a8;font-weight:700;color:#111;text-decoration:none;}
    .gap td{height:18px;border-left:none;border-right:none;background:#fff;color:#111;}
  </style></head><body><table><colgroup>`;
  html += colWidths.map(w=>`<col style="width:${w}px">`).join('');
  html += `</colgroup><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>`;
  rows.forEach((r)=>{
    const d=dateOnly(r.date);
    if(multipleDates && currentDate && d!==currentDate){
      html += `<tr class="gap"><td colspan="${headers.length}">&nbsp;</td></tr>`;
      seen=new Set();
    }
    if(d!==currentDate){
      currentDate=d;
      if(multipleDates) seen=new Set();
    }
    displayNo++;
    html += '<tr>';
    html += `<td class="center normal">${displayNo}</td>`;
    html += `<td class="center normal">${escapeHtml(normalize(r.date))}</td>`;
    html += `<td class="center normal">${escapeHtml(normalize(r.time))}</td>`;
    html += `<td class="normal">${escapeHtml(normalize(r.place))}</td>`;
    html += `<td class="normal">${escapeHtml(normalize(r.arrangement))}</td>`;
    for(let i=0;i<maxPlayers;i++){
      const p=normalize(r.players[i]);
      const key=performerKey(p);
      let cls='';
      if(key){
        if(seen.has(key)) cls = (cls ? cls + ' ' : '') + 'dup';
        else seen.add(key);
      }
      html += cls ? `<td class="${cls}">${escapeHtml(p)}</td>` : `<td class="normal">${escapeHtml(p)}</td>`;
    }
    html += '</tr>';
  });
  html += '</tbody></table></body></html>';
  const blob = new Blob(['\ufeff', html], {type:'application/vnd.ms-excel;charset=utf-8;'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=`BNS_당일진행표_${new Date().toISOString().slice(0,10)}.xls`;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
