/* ─── Firebase 初期化 ─── */
const firebaseConfig={
  apiKey:"AIzaSyB312nvOuATNd7i0hHOEnPcBddkR2grA-s",
  authDomain:"gantt-chart-92d03.firebaseapp.com",
  databaseURL:"https://gantt-chart-92d03-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId:"gantt-chart-92d03",
  storageBucket:"gantt-chart-92d03.firebasestorage.app",
  messagingSenderId:"268008972196",
  appId:"1:268008972196:web:2f5b6d4bbf3834e61d6b31"
};
firebase.initializeApp(firebaseConfig);
const db=firebase.database();
const tasksRef=db.ref('tasks');
const holidaysRef=db.ref('holidays');

const DAY_W=28;
let tasks=[], holidays=[], viewStart=new Date(), editId=null, syncing=false, pendingDelId=null;
let editSubtasks=[], editingSubIdx=-1;
let loadedCount=0;
let currentTab='main';
let filterText='';
let importPendingTasks=[];
let dragState=null;
let barEditorState={taskId:null,subId:null};

/* 展開状態はブラウザごとにローカル保持 */
let expandedIds=new Set(JSON.parse(localStorage.getItem('gantt_expanded')||'[]'));
function saveExpanded(){localStorage.setItem('gantt_expanded',JSON.stringify([...expandedIds]));}

/* ─── ヘルパー関数 ─── */
function getKakouClass(v){
  if(!v) return '';
  if(v==='吉田機工') return 'k-yoshida';
  if(v.includes('タガミ')) return 'k-tagami';
  return 'k-other';
}

function addDays(ds,n){
  const d=new Date(ds); d.setDate(d.getDate()+n); return d.toISOString().slice(0,10);
}

function daysBetween(s,e){
  return Math.round((new Date(e)-new Date(s))/86400000);
}

function resolveKakouSaki(selectId,otherId){
  const v=document.getElementById(selectId).value;
  if(v==='その他') return document.getElementById(otherId).value.trim()||'その他';
  return v;
}

function onKakouSakiChange(selectId,otherId){
  const v=document.getElementById(selectId).value;
  document.getElementById(otherId).style.display=v==='その他'?'block':'none';
}

/* ─── INIT ─── */
function init(){
  viewStart=new Date(); viewStart.setDate(1);

  db.ref('.info/connected').on('value',snap=>{
    const ok=snap.val();
    document.getElementById('connDot').className='conn-dot '+(ok?'online':'offline');
    document.getElementById('connLabel').textContent=ok?'同期中':'オフライン';
  });

  tasksRef.on('value',snap=>{
    const data=snap.val();
    if(data===null){
      loadSampleData();
    }else{
      tasks=Object.values(data).sort((a,b)=>(a._order||0)-(b._order||0));
      tasks.forEach(t=>{t.expanded=expandedIds.has(String(t.id));});
    }
    onLoaded();
  });

  holidaysRef.on('value',snap=>{
    const data=snap.val();
    holidays=data?Object.entries(data).map(([date,h])=>({date,name:h.name||''})):[];
    onLoaded();
    if(!document.getElementById('calModal').classList.contains('hidden'))renderHolidayList();
  });

  document.getElementById('addBtn').onclick=()=>openModal();
  document.getElementById('cancelBtn').onclick=closeModal;
  document.getElementById('saveBtn').onclick=saveTask;
  document.getElementById('addSubBtn').onclick=addSubtaskRow;
  document.getElementById('modal').onclick=e=>{if(e.target.id==='modal')closeModal();};
  document.getElementById('prevM').onclick=()=>{viewStart.setMonth(viewStart.getMonth()-1);render();};
  document.getElementById('nextM').onclick=()=>{viewStart.setMonth(viewStart.getMonth()+1);render();};
  document.getElementById('todayBtn').onclick=()=>{viewStart=new Date();viewStart.setDate(1);render();setTimeout(scrollToToday,80);};
  document.getElementById('calBtn').onclick=openCalModal;
  document.getElementById('closeCal').onclick=closeCalModal;
  document.getElementById('calModal').onclick=e=>{if(e.target.id==='calModal')closeCalModal();};
  document.getElementById('addHdBtn').onclick=addHoliday;
  document.getElementById('hdDate').onkeydown=e=>{if(e.key==='Enter')addHoliday();};
  document.getElementById('hdName').onkeydown=e=>{if(e.key==='Enter')addHoliday();};
  document.querySelectorAll('.cal-tab').forEach(btn=>{
    btn.onclick=()=>{
      document.querySelectorAll('.cal-tab').forEach(b=>b.classList.remove('active'));
      document.querySelectorAll('.cal-panel').forEach(p=>p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-'+btn.dataset.tab).classList.add('active');
    };
  });

  /* Excel import */
  document.getElementById('importBtn').onclick=()=>document.getElementById('fileInput').click();
  document.getElementById('fileInput').onchange=e=>{
    if(e.target.files[0])handleImportFile(e.target.files[0]);
    e.target.value='';
  };
  document.getElementById('importCancelBtn').onclick=()=>{
    document.getElementById('importModal').classList.add('hidden');
    importPendingTasks=[];
  };
  document.getElementById('importOkBtn').onclick=executeImport;
  document.getElementById('importModal').onclick=e=>{if(e.target.id==='importModal'){document.getElementById('importModal').classList.add('hidden');importPendingTasks=[];}};

  /* Excel export */
  document.getElementById('exportBtn').onclick=exportExcel;

  const L=document.getElementById('gLeft'),R=document.getElementById('gRight');
  L.addEventListener('scroll',()=>{if(syncing)return;syncing=true;R.scrollTop=L.scrollTop;syncing=false;});
  R.addEventListener('scroll',()=>{if(syncing)return;syncing=true;L.scrollTop=R.scrollTop;syncing=false;});

  document.addEventListener('mousemove',e=>{
    /* tooltip follow */
    const tip=document.getElementById('tip');
    if(!tip.classList.contains('hidden')){
      tip.style.left=Math.min(e.clientX+14,window.innerWidth-290)+'px';
      tip.style.top=Math.min(e.clientY+14,window.innerHeight-150)+'px';
    }
    /* drag handling */
    if(dragState){
      e.preventDefault();
      const dx=e.clientX-dragState.startX;
      const daysDelta=Math.round(dx/DAY_W);
      if(daysDelta!==dragState.lastDelta){
        dragState.lastDelta=daysDelta;
        let newStart=dragState.origStart,newEnd=dragState.origEnd;
        const mode=dragState.mode||'move';
        if(mode==='move'){
          newStart=addDays(dragState.origStart,daysDelta);
          newEnd=addDays(dragState.origEnd,daysDelta);
          /* clamp subtask to parent */
          if(dragState.subId!==''){
            if(newStart<dragState.parentStart){
              const shift=daysBetween(newStart,dragState.parentStart);
              newStart=addDays(newStart,shift);newEnd=addDays(newEnd,shift);
            }
            if(newEnd>dragState.parentEnd){
              const shift=daysBetween(dragState.parentEnd,newEnd);
              newStart=addDays(newStart,-shift);newEnd=addDays(newEnd,-shift);
            }
          }
        }else if(mode==='resize-start'){
          newStart=addDays(dragState.origStart,daysDelta);
          newEnd=dragState.origEnd;
          if(newStart>=newEnd) newStart=addDays(newEnd,-1);
          if(dragState.subId!==''&&newStart<dragState.parentStart) newStart=dragState.parentStart;
        }else if(mode==='resize-end'){
          newStart=dragState.origStart;
          newEnd=addDays(dragState.origEnd,daysDelta);
          if(newEnd<=newStart) newEnd=addDays(newStart,1);
          if(dragState.subId!==''&&newEnd>dragState.parentEnd) newEnd=dragState.parentEnd;
        }
        dragState.previewStart=newStart;
        dragState.previewEnd=newEnd;
        /* update bar DOM directly */
        const days=getDays();
        const vs=days[0];
        const bx=ds=>Math.round((new Date(ds).setHours(0,0,0,0)-vs.getTime())/86400000);
        const si=bx(newStart),ei=bx(newEnd)+1;
        const cs=Math.max(0,si),ce=Math.min(days.length,ei);
        if(ce>cs&&dragState.el){
          dragState.el.style.left=(cs*DAY_W)+'px';
          dragState.el.style.width=((ce-cs)*DAY_W)+'px';
        }
      }
    }
  });

  document.addEventListener('mouseup',e=>{
    if(dragState){
      hideTip();
      if(dragState.lastDelta!==0){
        /* 日付変化あり → 保存 */
        const t=tasks.find(x=>x.id===dragState.taskId);
        if(t){
          if(dragState.subId===''){t.start=dragState.previewStart;t.end=dragState.previewEnd;}
          else{const st=(t.subtasks||[]).find(x=>x.id===dragState.subId);if(st){st.start=dragState.previewStart;st.end=dragState.previewEnd;}}
          saveTasks();
        }
        dragJustCompleted=true;
        setTimeout(()=>{dragJustCompleted=false;},80);
      }
      /* lastDelta===0 のときは onclick が自然に発火してポップアップを開く */
      dragState=null;
    }
  });

  document.addEventListener('click',e=>{
    if(pendingDelId!==null&&!e.target.closest('.confirm-del')){pendingDelId=null;renderList();}
    if(!e.target.closest('#barEditor')) document.getElementById('barEditor').classList.add('hidden');
    if(!e.target.closest('#stPopup')) closeSubPopup();
    if(!e.target.closest('#taskPopup')) closeTaskPopup();
  });
}

function onLoaded(){
  loadedCount++;
  render();
  if(loadedCount>=2){
    document.getElementById('loadingOverlay').classList.add('hidden');
    setTimeout(scrollToToday,80);
  }
}

/* サンプルデータ（初回のみ） */
function loadSampleData(){
  const t=new Date(),fmt=d=>d.toISOString().slice(0,10);
  const add=(d,n)=>{const x=new Date(d);x.setDate(x.getDate()+n);return x;};
  const sample=[
    {id:Date.now()+1,name:'GD955',koban:'24R001',num:'GD955',owner:'田中',start:fmt(add(t,-5)),end:fmt(add(t,20)),status:'進行中',prog:35,note:'',kakouSaki:'吉田機工',subtasks:[
      {id:'s11',name:'アジャスター',owner:'田中',start:fmt(add(t,-5)),end:fmt(add(t,0)),status:'完了',prog:100,kakouSaki:'吉田機工',note:''},
      {id:'s12',name:'サークル',owner:'田中',start:fmt(add(t,1)),end:fmt(add(t,12)),status:'進行中',prog:40,kakouSaki:'タガミ・イーエクス',note:''},
      {id:'s13',name:'ブレード',owner:'田中',start:fmt(add(t,13)),end:fmt(add(t,20)),status:'未着手',prog:0,kakouSaki:'吉田機工',note:'品番: 8234001'},
    ]},
    {id:Date.now()+2,name:'K56',koban:'24R002',num:'K56',owner:'鈴木',start:fmt(add(t,3)),end:fmt(add(t,30)),status:'未着手',prog:0,note:'',kakouSaki:'吉田機工',subtasks:[
      {id:'s21',name:'アジャスター',owner:'鈴木',start:fmt(add(t,3)),end:fmt(add(t,18)),status:'未着手',prog:0,kakouSaki:'吉田機工',note:''},
      {id:'s22',name:'サークル',owner:'鈴木',start:fmt(add(t,5)),end:fmt(add(t,28)),status:'未着手',prog:0,kakouSaki:'タガミ・イーエクス',note:''},
    ]},
    {id:Date.now()+3,name:'特殊機',koban:'',num:'特殊機',owner:'佐藤',start:fmt(add(t,-15)),end:fmt(add(t,5)),status:'遅延',prog:60,note:'仕入れ先要確認',kakouSaki:'',subtasks:[]},
  ];
  const obj={};
  sample.forEach((t,i)=>{obj[String(t.id)]={...t,_order:i};});
  tasksRef.set(obj);
}

/* ─── Firebase 書き込み ─── */
function saveTasks(){
  const obj={};
  tasks.forEach((t,i)=>{
    const d={...t,_order:i};
    delete d.expanded;
    obj[String(t.id)]=d;
  });
  tasksRef.set(obj);
}
function saveHolidays(){
  const obj={};
  holidays.forEach(h=>{obj[h.date]={name:h.name||''};});
  holidaysRef.set(obj);
}

/* ─── タブ切替 ─── */
function switchTab(tab){
  currentTab=tab;
  document.getElementById('tabMain').classList.toggle('active',tab==='main');
  document.getElementById('tabMachining').classList.toggle('active',tab==='machining');
  document.getElementById('machiningInfo').classList.toggle('hidden',tab!=='machining');
  render();
}

/* ─── フィルター ─── */
function onFilterInput(val){
  filterText=val.trim();
  document.getElementById('filterClear').style.display=filterText?'flex':'none';
  render();
}
function clearFilter(){
  filterText='';
  document.getElementById('filterInput').value='';
  document.getElementById('filterClear').style.display='none';
  render();
}

function getDisplayTasks(){
  let list=tasks;
  if(filterText){
    const q=filterText.toLowerCase();
    list=list.filter(t=>
      t.name.toLowerCase().includes(q)||
      (t.num||'').toLowerCase().includes(q)||
      (t.koban||'').toLowerCase().includes(q)
    );
  }
  if(currentTab==='machining'){
    list=list.filter(t=>t.subtasks&&t.subtasks.some(st=>st.name.includes('機械加工')));
  } else {
    list=list.filter(t=>!t.machiningOnly);
  }
  return list;
}

/* ─── RENDER ─── */
function render(){
  const y=viewStart.getFullYear(),m=viewStart.getMonth()+1;
  document.getElementById('monthLbl').textContent=y+'年'+m+'月';
  renderSummary();renderList();renderTimeline();
}

function renderSummary(){
  const today=new Date();today.setHours(0,0,0,0);
  const list=getDisplayTasks();
  let act=0,don=0,ove=0;
  list.forEach(t=>{
    if(t.status==='進行中')act++;
    if(t.status==='完了')don++;
    const e=new Date(t.end);e.setHours(0,0,0,0);
    if(t.status!=='完了'&&e<today)ove++;
  });
  document.getElementById('s-all').textContent=list.length;
  document.getElementById('s-act').textContent=act;
  document.getElementById('s-don').textContent=don;
  document.getElementById('s-ove').textContent=ove;
}

function kakouBadgeHtml(v){
  if(!v) return '';
  const cls=getKakouClass(v);
  return `<span class="k-badge ${cls}">${esc(v)}</span>`;
}

function renderList(){
  const c=document.getElementById('taskRows');
  const list=getDisplayTasks();
  if(!list.length){
    c.innerHTML='<div class="empty-state"><i class="ti ti-clipboard-list" aria-hidden="true"></i><br>'+(filterText||currentTab==='machining'?'該当する案件がありません':'案件がありません<br>「案件追加」から始めましょう')+'</div>';
    return;
  }
  let html='';
  list.forEach(t=>{
    const hasSub=t.subtasks&&t.subtasks.length>0;
    const isPendingDel=pendingDelId===t.id;
    const kobanTxt=t.koban?`<span style="color:#5B89C8;margin-left:4px;">工番: ${esc(t.koban)}</span>`:'';
    const kBadge=kakouBadgeHtml(t.kakouSaki);
    html+=`<div class="t-row" data-id="${t.id}">
      <div class="t-cell" style="padding:0 4px;">
        ${hasSub?`<button class="expand-btn ${t.expanded?'open':''}" onclick="toggleExpand(event,${t.id})"><i class="ti ti-chevron-right"></i></button>`:'<span style="width:20px;display:inline-block"></span>'}
      </div>
      <div class="t-cell" onclick="openModal(${t.id})" style="cursor:pointer;">
        <div class="task-name">${esc(t.name)}</div>
        <div class="task-num">${esc(t.num||'—')}${kobanTxt}</div>
        ${kBadge?`<div style="margin-top:2px">${kBadge}</div>`:''}
      </div>
      <div class="t-cell" onclick="openModal(${t.id})" style="cursor:pointer"><span class="badge badge-${t.status}">${t.status}</span></div>
      <div class="t-cell" onclick="openModal(${t.id})" style="font-size:11px;color:#888;cursor:pointer">${esc(t.owner||'—')}</div>
      <div class="t-cell" onclick="openModal(${t.id})" style="cursor:pointer;font-size:11px;color:#555">${esc(t.stampNo||'—')}</div>
      <div class="t-cell"><div class="row-actions">
        ${isPendingDel
          ?`<button class="act-btn del confirm-del" onclick="confirmDel(event,${t.id})">削除確認</button>`
          :`<button class="act-btn del" onclick="requestDel(event,${t.id})" title="削除"><i class="ti ti-trash"></i></button>`}
      </div></div>
    </div>`;
    const stsToShow=currentTab==='machining'
      ?(t.expanded?(t.subtasks||[]).filter(st=>st.name.includes('機械加工')):[])
      :(t.expanded?(t.subtasks||[]):[]);
    if(stsToShow.length){
      stsToShow.forEach(st=>{
        const isMachining=st.name.includes('機械加工');
        const machiningCls=isMachining?' machining-row':'';
        const stKBadge=kakouBadgeHtml(st.kakouSaki);
        html+=`<div class="st-row${machiningCls}">
          <div class="t-cell"></div>
          <div class="t-cell" onclick="openModal(${t.id})" style="cursor:pointer;">
            <div class="st-indent">
              ${isMachining?'<i class="ti ti-tool" style="font-size:11px;color:#2DAA7A;flex-shrink:0"></i>':'<i class="ti ti-corner-down-right" style="font-size:11px;color:#bbb;flex-shrink:0"></i>'}
              <span class="st-name">${esc(st.name)}</span>
              ${stKBadge?`<span>${stKBadge}</span>`:''}
            </div>
            ${st.note?`<div class="st-note">${esc(st.note)}</div>`:''}
          </div>
          <div class="t-cell"><span class="badge badge-${st.status}" style="font-size:9px">${st.status}</span></div>
          <div class="t-cell" style="font-size:10px;color:#aaa">${esc(st.owner||'—')}</div>
          <div class="t-cell" style="font-size:10px;color:#777">${esc(st.koban||(t.koban?t.koban+'(引継)':'—'))}</div>
          <div class="t-cell"></div>
        </div>`;
      });
    }
  });
  c.innerHTML=html;
}

function toggleExpand(e,id){
  e.stopPropagation();
  const sid=String(id);
  if(expandedIds.has(sid))expandedIds.delete(sid);else expandedIds.add(sid);
  saveExpanded();
  tasks.forEach(t=>{t.expanded=expandedIds.has(String(t.id));});
  render();
}

function requestDel(e,id){
  e.stopPropagation();
  pendingDelId=id;renderList();
  setTimeout(()=>{if(pendingDelId===id){pendingDelId=null;renderList();}},3000);
}
function confirmDel(e,id){
  e.stopPropagation();
  tasks=tasks.filter(t=>t.id!==id);pendingDelId=null;saveTasks();render();
}

/* ─── TIMELINE ─── */
const holSet=()=>new Set(holidays.map(h=>h.date));

function getDays(){
  const s=new Date(viewStart);s.setDate(1);
  const end=new Date(s);end.setMonth(end.getMonth()+3);
  const days=[],c=new Date(s);
  while(c<end){days.push(new Date(c));c.setDate(c.getDate()+1);}
  return days;
}

function renderTimeline(){
  const days=getDays(),totalW=days.length*DAY_W;
  const today=new Date();today.setHours(0,0,0,0);
  const hs=holSet();
  const list=getDisplayTasks();

  const months={};
  days.forEach(d=>{
    const k=d.getFullYear()+'-'+d.getMonth();
    if(!months[k])months[k]={label:(d.getMonth()+1)+'月',count:0};
    months[k].count++;
  });
  const mRow=Object.values(months).map(m=>`<div class="month-label" style="width:${m.count*DAY_W}px">${m.label}</div>`).join('');

  const toLocal=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  const dRow=days.map(d=>{
    const isToday=d.getTime()===today.getTime();
    const dow=d.getDay();
    const ds=toLocal(d);
    const isHol=hs.has(ds);
    let cls='day-cell';
    if(isToday) cls+=' today';
    else if(isHol||dow===0) cls+=' sun';
    else if(dow===6) cls+=' sat';
    return `<div class="${cls}" style="width:${DAY_W}px">${d.getDate()}</div>`;
  }).join('');

  document.getElementById('tlHead').innerHTML=
    `<div class="month-row" style="width:${totalW}px">${mRow}</div>`+
    `<div class="day-row" style="width:${totalW}px">${dRow}</div>`;

  const ba=document.getElementById('barsArea');
  ba.style.width=totalW+'px';

  if(!list.length){ba.style.minHeight='80px';ba.innerHTML='';return;}

  let colBgs='';
  days.forEach((d,i)=>{
    const dow=d.getDay();
    const ds=toLocal(d);
    const isHol=hs.has(ds);
    if(dow===6) colBgs+=`<div class="col-bg col-sat" style="left:${i*DAY_W}px;width:${DAY_W}px"></div>`;
    else if(dow===0||isHol) colBgs+=`<div class="col-bg col-sun" style="left:${i*DAY_W}px;width:${DAY_W}px"></div>`;
  });

  let overlays='';
  days.forEach((d,i)=>{if(d.getDay()===1) overlays+=`<div class="grid-line" style="left:${i*DAY_W}px;z-index:1"></div>`;});
  const ti=days.findIndex(d=>d.getTime()===today.getTime());
  if(ti>=0) overlays+=`<div class="today-line" style="left:${ti*DAY_W+DAY_W/2}px"></div>`;

  const vs=days[0];
  function barX(ds){return Math.round((new Date(ds).setHours(0,0,0,0)-vs.getTime())/86400000);}
  function barW(s,e){return Math.max(1,barX(e)-barX(s)+1);}
  function clip(si,wi){
    const ei=si+wi;
    if(ei<=0||si>=days.length)return null;
    const cs=Math.max(0,si),ce=Math.min(days.length,ei);
    return{left:cs*DAY_W,width:(ce-cs)*DAY_W};
  }

  let bars='',top=0;
  list.forEach(t=>{
    const ROW=40;
    bars+=`<div class="row-divider" style="top:${top+ROW}px"></div>`;
    const c=clip(barX(t.start),barW(t.start,t.end));
    if(c){
      const lbl=esc(t.name);
      bars+=`<div id="bar-${t.id}" class="gantt-bar bar-parent bar-${t.status}" style="left:${c.left}px;width:${c.width}px;top:${top+10}px" onclick="barClick(event,${t.id},'')" onmousedown="barDown(event,${t.id},'')" onmousemove="updateBarCursor(event,this)" onmouseleave="hideTip();this.style.cursor='grab'" onmouseenter="showTip(event,${t.id})">${lbl}</div>`;
    }
    top+=ROW;

    const subsToShow=currentTab==='machining'
      ?(t.expanded?(t.subtasks||[]).filter(st=>st.name.includes('機械加工')):[])
      :(t.expanded?(t.subtasks||[]):[]);

    subsToShow.forEach(st=>{
      const SROW=34;
      bars+=`<div class="row-divider" style="top:${top+SROW}px"></div>`;
      const sc=clip(barX(st.start),barW(st.start,st.end));
      if(sc){
        const isMachining=st.name.includes('機械加工');
        const barCls=isMachining?'bar-機械加工':'bar-'+st.status;
        bars+=`<div id="bar-${t.id}-${st.id}" class="gantt-bar bar-sub ${barCls}" style="left:${sc.left}px;width:${sc.width}px;top:${top+10}px" onclick="barClick(event,${t.id},'${st.id}')" onmousedown="barDown(event,${t.id},'${st.id}')" onmousemove="updateBarCursor(event,this)" onmouseleave="hideTip();this.style.cursor='grab'" onmouseenter="showSubTip(event,${t.id},'${st.id}')">${esc(st.name)}</div>`;
      }
      top+=SROW;
    });
  });

  ba.style.minHeight=top+'px';
  ba.innerHTML=`<div style="position:relative;min-height:${top}px">${colBgs}${overlays}${bars}</div>`;
}

function scrollToToday(){
  const days=getDays(),today=new Date();today.setHours(0,0,0,0);
  const idx=days.findIndex(d=>d.getTime()===today.getTime());
  if(idx>=0)document.getElementById('gRight').scrollLeft=Math.max(0,idx*DAY_W-120);
}

/* ─── Drag to move/resize bars ─── */
let dragJustCompleted=false;

function barClick(e,taskId,subId){
  if(dragJustCompleted)return;
  e.stopPropagation();
  if(subId!=='') openSubPopup(taskId,subId,e.clientX,e.clientY);
  else openTaskPopup(taskId,e.clientX,e.clientY);
}

function barDown(e,taskId,subId){
  hideTip();
  const t=tasks.find(x=>x.id===taskId);
  if(!t) return;
  let origStart,origEnd,parentStart,parentEnd;
  if(subId===''){
    origStart=t.start; origEnd=t.end;
    parentStart=null; parentEnd=null;
  }else{
    const st=(t.subtasks||[]).find(x=>x.id===subId);
    if(!st) return;
    origStart=st.start; origEnd=st.end;
    parentStart=t.start; parentEnd=t.end;
  }
  const barId=subId===''?`bar-${taskId}`:`bar-${taskId}-${subId}`;
  const el=document.getElementById(barId);
  /* 端10px以内ならリサイズ、それ以外は移動 */
  const rect=e.currentTarget.getBoundingClientRect();
  const relX=e.clientX-rect.left;
  let mode='move';
  if(relX<=10) mode='resize-start';
  else if(relX>=rect.width-10) mode='resize-end';
  dragState={
    el,taskId,subId,mode,
    startX:e.clientX,
    origStart,origEnd,
    parentStart,parentEnd,
    lastDelta:0,
    previewStart:origStart,
    previewEnd:origEnd
  };
}

function updateBarCursor(e,el){
  if(dragState) return;
  const rect=el.getBoundingClientRect();
  const relX=e.clientX-rect.left;
  el.style.cursor=(relX<=10||relX>=rect.width-10)?'ew-resize':'grab';
}

/* ─── Bar click inline editor ─── */
function showBarEditor(taskId,subId,cx,cy){
  const t=tasks.find(x=>x.id===taskId);
  if(!t) return;
  barEditorState={taskId,subId};
  const ed=document.getElementById('barEditor');
  if(subId===''){
    document.getElementById('beTitle').textContent=t.name;
    document.getElementById('beStart').value=t.start;
    document.getElementById('beEnd').value=t.end;
    document.getElementById('beHint').textContent='';
  }else{
    const st=(t.subtasks||[]).find(x=>x.id===subId);
    if(!st){ed.classList.add('hidden');return;}
    document.getElementById('beTitle').textContent='↳ '+st.name;
    document.getElementById('beStart').value=st.start;
    document.getElementById('beEnd').value=st.end;
    document.getElementById('beHint').textContent='親タスク範囲: '+t.start+' 〜 '+t.end;
  }
  ed.classList.remove('hidden');
  const w=230,h=180;
  let left=cx+8,top=cy+8;
  if(left+w>window.innerWidth) left=cx-w-8;
  if(top+h>window.innerHeight) top=cy-h-8;
  ed.style.left=Math.max(4,left)+'px';
  ed.style.top=Math.max(4,top)+'px';
}

function saveBarEdit(){
  const newStart=document.getElementById('beStart').value;
  const newEnd=document.getElementById('beEnd').value;
  if(!newStart||!newEnd){alert('開始日と終了日を入力してください');return;}
  if(newStart>newEnd){alert('終了日は開始日以降にしてください');return;}
  const {taskId,subId}=barEditorState;
  const t=tasks.find(x=>x.id===taskId);
  if(!t){document.getElementById('barEditor').classList.add('hidden');return;}
  if(subId===''){
    t.start=newStart; t.end=newEnd;
  }else{
    const st=(t.subtasks||[]).find(x=>x.id===subId);
    if(st){
      let s=newStart,en=newEnd;
      if(s<t.start) s=t.start;
      if(en>t.end) en=t.end;
      if(s>en) en=s;
      st.start=s; st.end=en;
    }
  }
  saveTasks();
  document.getElementById('barEditor').classList.add('hidden');
}

function cancelBarEdit(){
  document.getElementById('barEditor').classList.add('hidden');
}

/* ─── サブタスク編集ポップアップ ─── */
let subPopupState=null;

function openSubPopup(taskId,subId,cx,cy){
  closeTaskPopup();
  const t=tasks.find(x=>x.id===taskId);
  if(!t)return;
  const st=(t.subtasks||[]).find(x=>x.id===subId);
  if(!st)return;
  subPopupState={taskId,subId};

  document.getElementById('spTitle').textContent=st.name||'工程を編集';
  document.getElementById('spName').value=st.name||'';
  document.getElementById('spStart').value=st.start||'';
  document.getElementById('spEnd').value=st.end||'';
  document.getElementById('spStart').min=t.start;
  document.getElementById('spStart').max=t.end;
  document.getElementById('spEnd').min=t.start;
  document.getElementById('spEnd').max=t.end;
  document.getElementById('spHint').textContent='親タスク範囲: '+t.start+' 〜 '+t.end;
  document.getElementById('spStatus').value=st.status||'未着手';
  document.getElementById('spOwner').value=st.owner||'';
  document.getElementById('spNote').value=st.note||'';

  const ks=st.kakouSaki||'';
  const known=['','吉田機工','タガミ・イーエクス'];
  if(known.includes(ks)){
    document.getElementById('spKakouSaki').value=ks;
    document.getElementById('spKakouSakiOther').style.display='none';
    document.getElementById('spKakouSakiOther').value='';
  }else{
    document.getElementById('spKakouSaki').value='その他';
    document.getElementById('spKakouSakiOther').value=ks;
    document.getElementById('spKakouSakiOther').style.display='block';
  }

  const pop=document.getElementById('stPopup');
  pop.classList.remove('hidden');
  const W=300,H=420;
  let left=cx+10,top=cy+10;
  if(left+W>window.innerWidth-8) left=cx-W-10;
  if(top+H>window.innerHeight-8) top=cy-H-10;
  pop.style.left=Math.max(4,left)+'px';
  pop.style.top=Math.max(4,top)+'px';
}

function saveSubPopup(){
  if(!subPopupState)return;
  const name=document.getElementById('spName').value.trim();
  const start=document.getElementById('spStart').value;
  const end=document.getElementById('spEnd').value;
  if(!name){document.getElementById('spName').focus();return;}
  if(!start||!end){alert('開始日と終了日を入力してください');return;}
  if(start>end){alert('終了日は開始日以降にしてください');return;}

  const t=tasks.find(x=>x.id===subPopupState.taskId);
  if(!t){closeSubPopup();return;}
  const st=(t.subtasks||[]).find(x=>x.id===subPopupState.subId);
  if(!st){closeSubPopup();return;}

  /* 親範囲でクランプ */
  const cs=start<t.start?t.start:start;
  const ce=end>t.end?t.end:end;

  st.name=name;
  st.start=cs;
  st.end=ce;
  st.status=document.getElementById('spStatus').value;
  st.owner=document.getElementById('spOwner').value.trim();
  st.note=document.getElementById('spNote').value.trim();
  st.kakouSaki=resolveKakouSaki('spKakouSaki','spKakouSakiOther');

  saveTasks();
  closeSubPopup();
}

function closeSubPopup(){
  document.getElementById('stPopup').classList.add('hidden');
  subPopupState=null;
}

/* ─── 親タスク編集ポップアップ ─── */
let taskPopupState=null;

function openTaskPopup(taskId,cx,cy){
  closeSubPopup();
  const t=tasks.find(x=>x.id===taskId);if(!t)return;
  taskPopupState={taskId};
  document.getElementById('tpTitle').textContent=t.name;
  document.getElementById('tpName').value=t.name;
  document.getElementById('tpKoban').value=t.koban||'';
  document.getElementById('tpStampNo').value=t.stampNo||'';
  document.getElementById('tpNum').value=t.num||'';
  document.getElementById('tpOwner').value=t.owner||'';
  document.getElementById('tpStart').value=t.start;
  document.getElementById('tpEnd').value=t.end;
  document.getElementById('tpStatus').value=t.status;
  document.getElementById('tpNote').value=t.note||'';
  const ks=t.kakouSaki||'';
  const known=['','吉田機工','タガミ・イーエクス'];
  if(known.includes(ks)){
    document.getElementById('tpKakouSaki').value=ks;
    document.getElementById('tpKakouSakiOther').style.display='none';
    document.getElementById('tpKakouSakiOther').value='';
  }else{
    document.getElementById('tpKakouSaki').value='その他';
    document.getElementById('tpKakouSakiOther').value=ks;
    document.getElementById('tpKakouSakiOther').style.display='block';
  }
  const pop=document.getElementById('taskPopup');
  pop.classList.remove('hidden');
  const W=300,H=520;
  let left=cx+10,top=cy+10;
  if(left+W>window.innerWidth-8) left=cx-W-10;
  if(top+H>window.innerHeight-8) top=cy-H-10;
  pop.style.left=Math.max(4,left)+'px';
  pop.style.top=Math.max(4,top)+'px';
}

function saveTaskPopup(){
  if(!taskPopupState)return;
  const name=document.getElementById('tpName').value.trim();
  const start=document.getElementById('tpStart').value;
  const end=document.getElementById('tpEnd').value;
  if(!name){document.getElementById('tpName').focus();return;}
  if(!start||!end){alert('開始日と納期を入力してください');return;}
  if(start>end){alert('納期は開始日以降にしてください');return;}
  const t=tasks.find(x=>x.id===taskPopupState.taskId);
  if(!t){closeTaskPopup();return;}
  t.name=name;
  t.koban=document.getElementById('tpKoban').value.trim();
  t.stampNo=document.getElementById('tpStampNo').value.trim();
  t.num=document.getElementById('tpNum').value.trim();
  t.owner=document.getElementById('tpOwner').value.trim();
  t.start=start;t.end=end;
  t.status=document.getElementById('tpStatus').value;
  t.note=document.getElementById('tpNote').value.trim();
  t.kakouSaki=resolveKakouSaki('tpKakouSaki','tpKakouSakiOther');
  saveTasks();closeTaskPopup();
}

function closeTaskPopup(){
  document.getElementById('taskPopup').classList.add('hidden');
  taskPopupState=null;
}

/* ─── 案件モーダル ─── */
function openModal(id){
  editId=id||null;
  document.getElementById('modal').classList.remove('hidden');
  document.getElementById('subtaskSection').style.display=id?'block':'none';
  if(id){
    const t=tasks.find(x=>x.id===id);if(!t)return;
    editSubtasks=JSON.parse(JSON.stringify(t.subtasks||[]));
    document.getElementById('mTitle').textContent='案件を編集';
    document.getElementById('mSub').textContent='';
    setFields(t);renderSubtaskList();
  }else{
    editSubtasks=[];
    document.getElementById('mTitle').textContent='案件を追加';
    document.getElementById('mSub').textContent='';
    clearFields();
  }
  document.getElementById('fName').focus();
}

function setFields(t){
  document.getElementById('fName').value=t.name;
  document.getElementById('fKoban').value=t.koban||'';
  document.getElementById('fStampNo').value=t.stampNo||'';
  document.getElementById('fNum').value=t.num||'';
  document.getElementById('fOwner').value=t.owner||'';
  document.getElementById('fStart').value=t.start;
  document.getElementById('fEnd').value=t.end;
  document.getElementById('fStatus').value=t.status;
  document.getElementById('fNote').value=t.note||'';
  /* 加工先 */
  const ks=t.kakouSaki||'';
  const knownVals=['','吉田機工','タガミ・イーエクス'];
  if(knownVals.includes(ks)){
    document.getElementById('fKakouSaki').value=ks;
    document.getElementById('fKakouSakiOther').value='';
    document.getElementById('fKakouSakiOther').style.display='none';
  }else{
    document.getElementById('fKakouSaki').value='その他';
    document.getElementById('fKakouSakiOther').value=ks;
    document.getElementById('fKakouSakiOther').style.display='block';
  }
}
function clearFields(){
  ['fName','fKoban','fStampNo','fNum','fOwner','fNote'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('fStart').value=new Date().toISOString().slice(0,10);
  document.getElementById('fEnd').value='';
  document.getElementById('fStatus').value='未着手';
  document.getElementById('fKakouSaki').value='';
  document.getElementById('fKakouSakiOther').value='';
  document.getElementById('fKakouSakiOther').style.display='none';
}

function renderSubtaskList(){
  const list=document.getElementById('subtaskList');
  if(!editSubtasks.length){list.innerHTML='<div style="font-size:11px;color:#bbb;padding:4px 0">工程なし</div>';return;}
  list.innerHTML=editSubtasks.map((st,i)=>{
    if(i===editingSubIdx){
      return `<div class="st-edit-form">
        <div class="st-edit-row">
          <input type="text" id="se-name" value="${esc(st.name)}" placeholder="工程名">
          <input type="date" id="se-start" value="${st.start}">
          <input type="date" id="se-end" value="${st.end}">
        </div>
        <div class="st-edit-row2">
          <select id="se-status">
            ${['未着手','進行中','完了','遅延','保留'].map(s=>`<option value="${s}"${st.status===s?' selected':''}>${s}</option>`).join('')}
          </select>
          <select id="se-koban-mode" onchange="document.getElementById('se-koban').style.display=this.value==='custom'?'block':'none'">
            <option value="inherit"${!st.koban?' selected':''}>工番：親タスク引継ぎ</option>
            <option value="custom"${st.koban?' selected':''}>工番：別途入力</option>
          </select>
          <input type="text" id="se-koban" placeholder="工番" value="${esc(st.koban||'')}" style="display:${st.koban?'block':'none'}">
        </div>
        <div class="st-edit-row3">
          <select id="se-kakou" onchange="onKakouSakiChange('se-kakou','se-kakou-other')">
            <option value="">加工先未設定</option>
            <option value="吉田機工"${(st.kakouSaki==='吉田機工')?' selected':''}>吉田機工</option>
            <option value="タガミ・イーエクス"${(st.kakouSaki==='タガミ・イーエクス')?' selected':''}>タガミ・イーエクス</option>
            <option value="その他"${(st.kakouSaki&&st.kakouSaki!=='吉田機工'&&st.kakouSaki!=='タガミ・イーエクス')?' selected':''}>その他</option>
          </select>
          <input type="text" id="se-kakou-other" placeholder="加工先" value="${(st.kakouSaki&&st.kakouSaki!=='吉田機工'&&st.kakouSaki!=='タガミ・イーエクス')?esc(st.kakouSaki):''}" style="display:${(st.kakouSaki&&st.kakouSaki!=='吉田機工'&&st.kakouSaki!=='タガミ・イーエクス')?'block':'none'}">
          <input type="text" id="se-note" placeholder="備考" value="${esc(st.note||'')}">
        </div>
        <div class="st-edit-row3">
          <input type="text" id="se-stamp" placeholder="刻印No." value="${esc(st.stampNo||'')}" style="flex:1;">
        </div>
        <div class="st-edit-actions">
          <button onclick="cancelEditSub()">キャンセル</button>
          <button class="st-save-btn" onclick="saveEditSub(${i})"><i class="ti ti-check"></i> 保存</button>
        </div>
      </div>`;
    }
    const isMachining=st.name.includes('機械加工');
    const stKBadge=kakouBadgeHtml(st.kakouSaki);
    const parentKoban=document.getElementById('fKoban').value.trim();
    const dispKoban=st.koban||(parentKoban?parentKoban+'(引継)':'');
    return `<div class="st-item${isMachining?' machining-item':''}">
      ${isMachining?'<i class="ti ti-tool" style="color:#2DAA7A;font-size:13px;flex-shrink:0"></i>':''}
      <div class="st-item-name">${esc(st.name)}${stKBadge?' '+stKBadge:''}</div>
      <div class="st-item-info">${st.start} 〜 ${st.end}</div>
      ${dispKoban?`<span class="st-note">工番: ${esc(dispKoban)}</span>`:''}
      ${st.stampNo?`<span class="st-note">刻印: ${esc(st.stampNo)}</span>`:''}
      ${st.note?`<span class="st-note">${esc(st.note)}</span>`:''}
      <span class="badge badge-${st.status}" style="font-size:9px">${st.status}</span>
      <button class="st-item-edit-btn" onclick="startEditSub(${i})" title="編集"><i class="ti ti-pencil"></i></button>
      <button class="st-item-del" onclick="removeSubtask(${i})" title="削除"><i class="ti ti-x"></i></button>
    </div>`;
  }).join('');
}
function startEditSub(i){editingSubIdx=i;renderSubtaskList();}
function cancelEditSub(){editingSubIdx=-1;renderSubtaskList();}
function saveEditSub(i){
  const name=document.getElementById('se-name').value.trim();
  const start=document.getElementById('se-start').value;
  const end=document.getElementById('se-end').value;
  if(!name){document.getElementById('se-name').focus();return;}
  if(!start||!end){alert('開始日と終了日を入力してください');return;}
  if(start>end){alert('終了日は開始日以降にしてください');return;}
  const kakouVal=document.getElementById('se-kakou').value;
  const kakouSaki=kakouVal==='その他'?(document.getElementById('se-kakou-other').value.trim()||'その他'):kakouVal;
  const note=document.getElementById('se-note').value.trim();
  const kobanMode=document.getElementById('se-koban-mode').value;
  const koban=kobanMode==='custom'?document.getElementById('se-koban').value.trim():'';
  const stampNo=document.getElementById('se-stamp').value.trim();
  editSubtasks[i]={...editSubtasks[i],name,start,end,status:document.getElementById('se-status').value,prog:0,kakouSaki,note,koban,stampNo};
  editingSubIdx=-1;renderSubtaskList();
}
function addSubtaskRow(){
  const name=document.getElementById('stName').value.trim();
  const start=document.getElementById('stStart').value,end=document.getElementById('stEnd').value;
  if(!name){document.getElementById('stName').focus();return;}
  if(!start||!end){alert('開始日と終了日を入力してください');return;}
  if(start>end){alert('終了日は開始日以降にしてください');return;}
  const kakouSaki=resolveKakouSaki('stKakouSaki','stKakouSakiOther');
  const note=document.getElementById('stNote').value.trim();
  const kobanMode=document.getElementById('stKobanMode').value;
  const koban=kobanMode==='custom'?document.getElementById('stKoban').value.trim():'';
  const stampNo=document.getElementById('stStampNo').value.trim();
  editSubtasks.push({id:'s'+Date.now(),name,owner:'',start,end,status:'未着手',prog:0,kakouSaki,note,koban,stampNo});
  document.getElementById('stName').value='';
  document.getElementById('stStart').value='';
  document.getElementById('stEnd').value='';
  document.getElementById('stKakouSaki').value='';
  document.getElementById('stKakouSakiOther').value='';
  document.getElementById('stKakouSakiOther').style.display='none';
  document.getElementById('stNote').value='';
  document.getElementById('stKobanMode').value='inherit';
  document.getElementById('stKoban').value='';
  document.getElementById('stKoban').style.display='none';
  document.getElementById('stStampNo').value='';
  renderSubtaskList();
  document.getElementById('stName').focus();
}
function removeSubtask(i){editSubtasks.splice(i,1);renderSubtaskList();}
function closeModal(){document.getElementById('modal').classList.add('hidden');editId=null;editSubtasks=[];editingSubIdx=-1;}
function saveTask(){
  const name=document.getElementById('fName').value.trim();
  const start=document.getElementById('fStart').value,end=document.getElementById('fEnd').value;
  if(!name){alert('案件名を入力してください');return;}
  if(!start){alert('開始日を入力してください');return;}
  if(!end){alert('納期を入力してください');return;}
  if(start>end){alert('納期は開始日以降に設定してください');return;}
  const kakouSaki=resolveKakouSaki('fKakouSaki','fKakouSakiOther');
  const d={
    name,
    koban:document.getElementById('fKoban').value.trim(),
    stampNo:document.getElementById('fStampNo').value.trim(),
    num:document.getElementById('fNum').value.trim(),
    owner:document.getElementById('fOwner').value.trim(),
    start,end,
    status:document.getElementById('fStatus').value,
    prog:0,
    note:document.getElementById('fNote').value.trim(),
    kakouSaki,
    subtasks:editSubtasks
  };
  if(editId){const i=tasks.findIndex(t=>t.id===editId);if(i>=0)tasks[i]={...tasks[i],...d,expanded:tasks[i].expanded};}
  else{d.id=Date.now();d.expanded=false;tasks.push(d);}
  saveTasks();closeModal();
}

/* ─── 休日モーダル ─── */
function openCalModal(){
  renderHolidayList();
  document.getElementById('calModal').classList.remove('hidden');
}
function closeCalModal(){document.getElementById('calModal').classList.add('hidden');}

function renderHolidayList(){
  const list=document.getElementById('holidayList');
  const sorted=[...holidays].sort((a,b)=>a.date.localeCompare(b.date));
  if(!sorted.length){list.innerHTML='<div class="hd-empty">登録済みの休日はありません</div>';return;}
  list.innerHTML=sorted.map(h=>{
    const d=new Date(h.date);
    const dow=['日','月','火','水','木','金','土'][d.getDay()];
    return `<div class="hd-item">
      <div class="hd-date">${h.date}（${dow}）</div>
      <div class="hd-name">${esc(h.name||'')}</div>
      <button class="hd-del" onclick="removeHoliday('${h.date}')"><i class="ti ti-x"></i></button>
    </div>`;
  }).join('');
}

function addHoliday(){
  const date=document.getElementById('hdDate').value;
  const name=document.getElementById('hdName').value.trim();
  if(!date){document.getElementById('hdDate').focus();return;}
  if(holidays.find(h=>h.date===date)){alert('その日付はすでに登録されています');return;}
  holidays.push({date,name});
  saveHolidays();
  document.getElementById('hdDate').value='';
  document.getElementById('hdName').value='';
}
function removeHoliday(date){
  holidays=holidays.filter(h=>h.date!==date);
  saveHolidays();
}

/* ─── Excel インポート（仕様名を親タスクに） ─── */
function parseJaDate(str){
  if(!str) return null;
  str=String(str).trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  const m=str.match(/^(\d{1,2})\/(\d{1,2})$/);
  if(!m) return null;
  const mon=parseInt(m[1]),day=parseInt(m[2]);
  if(mon<1||mon>12||day<1||day>31) return null;
  const now=new Date();
  const yr=mon<now.getMonth()+1?now.getFullYear()+1:now.getFullYear();
  return `${yr}-${String(mon).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
}

/* 仕様名 → サブタスク品種マップ */
const SPEC_SUBTASK_MAP={
  'GD825':['アジャスター'],
  'GD955':['アジャスター','ドローバ','サークル','ブレード'],
  'K56':['アジャスター','ドローバ','サークル'],
  '特殊機':['大）マテハン','エ）エクステ','エ）BUCKET','エ）JAW']
};

/* デフォルト加工先ルール */
function getDefaultKakouSaki(specName,partType){
  if(specName==='GD955'&&partType==='サークル') return 'タガミ・イーエクス';
  if(specName==='K56'&&partType==='サークル') return 'タガミ・イーエクス';
  if(specName==='特殊機'&&partType==='大）マテハン') return 'タガミ・イーエクス';
  return '吉田機工';
}

function extractSpecName(s){
  if(!s) return null;
  for(const k of Object.keys(SPEC_SUBTASK_MAP)){
    if(s.includes(k)) return k;
  }
  return null;
}

/* 品番から品種を抽出（最初の数字列の前のテキスト） */
function extractPartType(pnum){
  if(!pnum) return null;
  const m=pnum.match(/^([^\d]+)/);
  return m?m[1].trim():null;
}

async function handleImportFile(file){
  try{
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array'});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const range=XLSX.utils.decode_range(ws['!ref']||'A1');

    const cell=(r,c)=>{
      const addr=XLSX.utils.encode_cell({r,c});
      const v=ws[addr];
      if(!v) return '';
      if(v.t==='n'&&v.z&&v.z.includes('/')){
        const dt=XLSX.SSF.parse_date_code(v.v);
        if(dt) return `${dt.m}/${dt.d}`;
      }
      return String(v.v||'').trim();
    };

    /* グループ: {specName: {partType: [{pnum,koban,rows:[]}]}} */
    const specGroups={};
    let lastSpec='',lastPnum='';

    for(let r=6;r<=range.e.r;r++){
      const rawSpec=cell(r,0)||lastSpec;
      const pnum=cell(r,1)||lastPnum;
      const koban=cell(r,2);
      if(!koban) continue;

      const komatsuDeadline=cell(r,4);
      const transportOut=cell(r,7);
      const transportIn=cell(r,8);

      if(cell(r,0)) lastSpec=cell(r,0);
      if(cell(r,1)) lastPnum=cell(r,1);

      const specName=extractSpecName(rawSpec)||rawSpec||'不明';
      const partType=extractPartType(pnum)||pnum||'不明';

      if(!specGroups[specName]) specGroups[specName]={};
      if(!specGroups[specName][partType]) specGroups[specName][partType]=[];
      specGroups[specName][partType].push({pnum,koban,komatsuDeadline,transportOut,transportIn});
    }

    importPendingTasks=[];
    const usedIds=new Set();
    function uniqueId(){let id;do{id=Date.now()+Math.floor(Math.random()*1e6);}while(usedIds.has(id));usedIds.add(id);return id;}

    for(const [specName,partGroups] of Object.entries(specGroups)){
      /* 親タスクの日程を全サブタスクから集計 */
      let parentStarts=[],parentEnds=[];
      const subtasks=[];

      for(const [partType,rows] of Object.entries(partGroups)){
        const kakouSaki=getDefaultKakouSaki(specName,partType);

        /* 納期ごとに1サブタスク */
        for(const row of rows){
          const s=parseJaDate(row.transportOut);
          const e=parseJaDate(row.transportIn);
          if(!s||!e) continue;
          parentStarts.push(s); parentEnds.push(e);
          subtasks.push({id:'s'+uniqueId(),name:'機械加工（'+partType+'）',owner:'',start:s,end:e,status:'未着手',prog:0,kakouSaki,note:row.pnum||''});
        }
      }

      if(!subtasks.length) continue;

      /* コマツ納期を親の終了日に */
      const allDeadlines=[];
      for(const rows of Object.values(partGroups)){
        rows.forEach(row=>{const d=parseJaDate(row.komatsuDeadline);if(d)allDeadlines.push(d);});
      }
      allDeadlines.sort();
      const parentStart=parentStarts.sort()[0];
      const parentEnd=allDeadlines.length?allDeadlines[allDeadlines.length-1]:parentEnds.sort().pop();

      importPendingTasks.push({
        id:uniqueId(),
        name:specName,
        koban:'',
        num:'',
        owner:'',
        start:parentStart,
        end:parentEnd,
        status:'未着手',
        prog:0,
        note:'',
        kakouSaki:'',
        subtasks,
        expanded:true,
        machiningOnly:true
      });
    }

    if(!importPendingTasks.length){alert('取込可能なデータが見つかりませんでした。\nExcelの形式を確認してください。');return;}

    document.getElementById('importCount').textContent=`${importPendingTasks.length}件の案件をインポートします`;
    const rows=importPendingTasks.map(t=>`<tr>
      <td>${esc(t.name)}</td>
      <td>${t.start}</td>
      <td>${t.end}</td>
      <td>${t.subtasks.length}工程</td>
      <td>${t.subtasks.map(st=>`${esc(st.name)}（${esc(st.kakouSaki)}）`).join(', ')}</td>
    </tr>`).join('');
    document.getElementById('importPreview').innerHTML=`<table>
      <thead><tr><th>仕様名</th><th>開始日</th><th>納期</th><th>工程数</th><th>工程・加工先</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
    document.getElementById('importModal').classList.remove('hidden');
  }catch(err){
    console.error(err);
    alert('Excelの読み込みに失敗しました。\n'+err.message);
  }
}

function executeImport(){
  if(!importPendingTasks.length) return;
  tasks=[...tasks,...importPendingTasks];
  importPendingTasks.forEach(t=>{if(t.expanded)expandedIds.add(String(t.id));});
  saveExpanded();
  saveTasks();
  importPendingTasks=[];
  document.getElementById('importModal').classList.add('hidden');
}

/* ─── Excel エクスポート ─── */
function exportExcel(){
  const wb=XLSX.utils.book_new();

  /* Sheet1: 案件一覧 */
  const taskData=[['案件名','工番','品番','担当者','開始日','納期','ステータス','進捗率','加工先','備考']];
  tasks.forEach(t=>{
    taskData.push([t.name,t.koban||'',t.num||'',t.owner||'',t.start,t.end,t.status,t.prog+'%',t.kakouSaki||'',t.note||'']);
    (t.subtasks||[]).forEach(st=>{
      taskData.push(['  └ '+st.name,'','',st.owner||'',st.start,st.end,st.status,st.prog+'%',st.kakouSaki||'',st.note||'']);
    });
  });
  const ws1=XLSX.utils.aoa_to_sheet(taskData);
  ws1['!cols']=[{wch:28},{wch:12},{wch:16},{wch:10},{wch:12},{wch:12},{wch:8},{wch:8},{wch:18},{wch:24}];
  XLSX.utils.book_append_sheet(wb,ws1,'案件一覧');

  /* Sheet2: ガントチャート（日付グリッド） */
  const days=getDays();
  const ganttHeader=['案件名','工番','品番','担当者','ステータス',...days.map(d=>(d.getMonth()+1)+'/'+d.getDate())];
  const ganttData=[ganttHeader];
  tasks.forEach(t=>{
    const row=[t.name,t.koban||'',t.num||'',t.owner||'',t.status];
    days.forEach(d=>{
      const ds=d.toISOString().slice(0,10);
      row.push(ds>=t.start&&ds<=t.end?'■':'');
    });
    ganttData.push(row);
    if(t.subtasks&&t.subtasks.length){
      t.subtasks.forEach(st=>{
        const srow=['  └ '+st.name,'','','',st.status];
        days.forEach(d=>{
          const ds=d.toISOString().slice(0,10);
          srow.push(ds>=st.start&&ds<=st.end?'▪':'');
        });
        ganttData.push(srow);
      });
    }
  });
  const ws2=XLSX.utils.aoa_to_sheet(ganttData);
  ws2['!cols']=[{wch:28},{wch:12},{wch:14},{wch:10},{wch:8},...days.map(()=>({wch:3}))];
  XLSX.utils.book_append_sheet(wb,ws2,'ガントチャート');

  /* Sheet3: 機械加工専用日程表 */
  const mHeader=['案件名','工番','品番','担当者','開始日','終了日','ステータス','工程数','加工先','備考'];
  const mData=[mHeader];
  tasks.forEach(t=>{
    if(!(t.subtasks&&t.subtasks.length)) return;
    const pStart=t.subtasks.map(s=>s.start).sort()[0]||'';
    const pEnd=t.subtasks.map(s=>s.end).sort().pop()||'';
    mData.push([t.name,t.koban||'',t.num||'',t.owner||'',pStart,pEnd,t.status,t.subtasks.length,t.kakouSaki||'',t.note||'']);
    t.subtasks.forEach(st=>{
      mData.push(['  └ '+st.name,'','','',st.start,st.end,st.status,'',st.kakouSaki||'',st.note||'']);
    });
  });
  const ws3=XLSX.utils.aoa_to_sheet(mData);
  ws3['!cols']=[{wch:28},{wch:12},{wch:14},{wch:10},{wch:14},{wch:14},{wch:8},{wch:8},{wch:18},{wch:24}];
  XLSX.utils.book_append_sheet(wb,ws3,'機械加工専用日程表');

  const today=new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb,`ガントチャート_${today}.xlsx`);
}

/* ─── TOOLTIP ─── */
function showTip(e,id){
  if(dragState&&dragState.moved) return;
  const t=tasks.find(x=>x.id===id);if(!t)return;
  const tip=document.getElementById('tip');tip.classList.remove('hidden');
  document.getElementById('ttName').textContent=t.name;
  document.getElementById('ttKoban').textContent=t.koban?'工番: '+t.koban:'';
  document.getElementById('ttNum').textContent=t.num?'品番: '+t.num:'';
  document.getElementById('ttDates').textContent='期間: '+t.start+' 〜 '+t.end;
  document.getElementById('ttOwner').textContent=t.owner?'担当: '+t.owner:'';
  document.getElementById('ttProg').textContent='ステータス: '+t.status;
  document.getElementById('ttKakou').textContent=t.kakouSaki?'加工先: '+t.kakouSaki:'';
  document.getElementById('ttNote').textContent=t.note?'備考: '+t.note:'';
}
function showSubTip(e,pid,sid){
  if(dragState&&dragState.moved) return;
  const t=tasks.find(x=>x.id===pid);if(!t)return;
  const st=t.subtasks.find(x=>x.id==sid);if(!st)return;
  const tip=document.getElementById('tip');tip.classList.remove('hidden');
  document.getElementById('ttName').textContent='↳ '+st.name;
  document.getElementById('ttKoban').textContent='';
  document.getElementById('ttNum').textContent='親案件: '+t.name;
  document.getElementById('ttDates').textContent='期間: '+st.start+' 〜 '+st.end;
  document.getElementById('ttOwner').textContent=st.owner?'担当: '+st.owner:'';
  document.getElementById('ttProg').textContent='ステータス: '+st.status;
  document.getElementById('ttKakou').textContent=st.kakouSaki?'加工先: '+st.kakouSaki:'';
  document.getElementById('ttNote').textContent=st.note?'備考: '+st.note:'';
}
function hideTip(){document.getElementById('tip').classList.add('hidden');}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

init();
