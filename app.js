
const D=window.__PANEL_DATA__;
const DIAS=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
const DIAS3=['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
const SVC_MAP={};
D.svc_catalog.forEach(s=>{
  SVC_MAP[s.id]={
    id:s.id,local:s.l,direccion:s.d,entrada:s.e,salida:s.s,
    entrada_h:s.eh,salida_h:s.sh,dias:s.di,dias_num:s.dn,
    horas_dia:s.hd,horas_semana:s.hs,n_dias:s.nd,
    solo_sobre44:s.o,coords:s.c,zona:s.z,
    n_personas_actuales:s.np
  };
});

// Normalize compact schedule fields
const SCHED={};
Object.entries(D.schedules||{}).forEach(([nombre,sched])=>{
  SCHED[nombre]={};
  Object.entries(sched).forEach(([dia,slots])=>{
    SCHED[nombre][dia]=slots.map(s=>({
      local:s.l,entrada:s.e,salida:s.s,horas:s.h,entrada_h:s.eh,salida_h:s.sh
    }));
  });
});
// Normalize persona_disponibilidad compact fields
const PDISP={};
Object.entries(D.persona_disponibilidad||{}).forEach(([n,info])=>{
  PDISP[n]={horas_falt:info.hf,zonas:info.z,disponibilidad:{}};
  Object.entries(info.dsp||{}).forEach(([dia,slots])=>{
    PDISP[n].disponibilidad[dia]=slots.map(s=>({desde:s.de,hasta:s.ha,horas:s.h}));
  });
});
// Local name → coords lookup from svc_catalog
const LOCAL_COORDS={};
D.svc_catalog.forEach(s=>{if(s.c)LOCAL_COORDS[s.l.trim().toUpperCase()]=s.c;});




let _maps={};

function showTab(n,btn){
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+n).classList.add('active');btn.classList.add('active');
  if(n==='personas') setTimeout(()=>initMaps(),100);
}
function showSub(n,btn){
  document.querySelectorAll('.subtab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.subtab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('subtab-'+n).classList.add('active');btn.classList.add('active');
  setTimeout(()=>initMaps(),100);
}
function fH(h){
  if(h==null||isNaN(h))return'—';
  const t=Math.round(h*60),hh=Math.floor(t/60),mm=t%60;
  return mm===0?hh+'h':hh+'h '+mm+'min';
}
function esc(s){return(s||'').replace(/'/g,"\'").replace(/"/g,'&quot;')}

// ── EXPORT ──
function doExport(tipo){
  if(tipo==='comercial'){
    const b64=D.XL_COM||'';
    const bin=atob(b64),buf=new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++)buf[i]=bin.charCodeAt(i);
    const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);
    a.download='Analisis_Comercial_por_Zona.xlsx';a.click();
    return;
  }
  const b64=tipo==='menos44'?D.XL_MENOS||'':D.XL_MAS||'';
  const bin=atob(b64),buf=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++)buf[i]=bin.charCodeAt(i);
  const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download=tipo==='menos44'?'Menos_44h_Servicios.xlsx':'Mas_44h_Servicios.xlsx';a.click();
}

// ── DASHBOARD ──
function rDash(){
  const r=D.resumen;
  document.getElementById('badge-m44').textContent=r.menos44;
  document.getElementById('badge-mas44').textContent=r.mas44;
  const rg=r.rangos_menos44||{};
  document.getElementById('stat-cards').innerHTML=[
    {v:r.total_personas,l:'Total personas',s:'Analizadas esta semana',b:'accent-bar'},
    {v:rg.menos39||0,l:'Menos de 39h',s:'Urgente — muy por debajo',b:'red-bar'},
    {v:(rg.entre39y40||0)+(rg.h40||0),l:'Entre 39h y 40h',s:'Por debajo del objetivo',b:'amber-bar'},
    {v:rg.entre41y43||0,l:'Entre 41h y 43h',s:'Cerca del objetivo',b:'blue-bar'},
    {v:r.exactas44,l:'Con 44h exactas',s:'Rutina completa',b:'green-bar'},
    {v:r.mas44,l:'Más de 44h',s:'Control informativo',b:'green-bar'},
    {v:r.excluidos,l:'Excluidos',s:'Admin, Vidrieros, Maldonado',b:'amber-bar'},
  ].map(c=>`<div class="stat-card"><div class="sc-label">${c.l}</div><div class="sc-value">${c.v}</div><div class="sc-sub">${c.s}</div><div class="sc-bar ${c.b}"></div></div>`).join('');
  const dist=D.distribucion,ORD=['Menos de 39h','39h','40h','41h a 43h','44h exactas','Más de 44h'];
  const COL={'Menos de 39h':'#dc2626','39h':'#ea580c','40h':'#d97706','41h a 43h':'#4f46e5','44h exactas':'#059669','Más de 44h':'#047857'};
  const mx=Math.max(...ORD.map(k=>dist[k]||0));
  document.getElementById('dist-chart').innerHTML=ORD.map(k=>{
    if(!dist[k])return'';
    const pct=Math.round(dist[k]/mx*100);
    return`<div class="dist-bar-row"><div class="dist-bar-lbl">${k}</div><div class="dist-bar-track"><div class="dist-bar-fill" style="width:${pct}%;background:${COL[k]}">${dist[k]}</div></div><div class="dist-bar-n">${dist[k]}</div></div>`;
  }).join('');
}

// ── BUILD SERVICE EXPAND ──
function buildSvcBtn(svcs,id){
  if(!svcs||!svcs.length)return'<span style="color:var(--muted);font-size:11px">—</span>';
  return`<button class="expand-btn" id="ebtn-${id}" onclick="toggleSvc('${id}')">▶ ${svcs.length} servicio${svcs.length!==1?'s':''}</button>`;
}
function toggleSvc(id){
  const row=document.getElementById('srow-'+id);
  const btn=document.getElementById('ebtn-'+id);
  const o=row.classList.toggle('open');
  btn.classList.toggle('open',o);
  btn.textContent=o?'▼ Ocultar':'▶ Ver servicios';
}

// ── MENOS 44 ──
function rM44(){
  const d=D.menos44_enriched;
  document.getElementById('m44-cnt').textContent=`${d.length} personas`;
  let html='';
  d.forEach((p,i)=>{
    const rid='m44-'+i;
    const hc=p.horas<30?'chip-red':p.horas<40?'chip-amber':'chip-blue';
    const fc=p.horas_falt>10?'chip-red':p.horas_falt>5?'chip-amber':'chip-green';
    const dh=p.dias.map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join('');
    const zh=p.zonas.map(z=>`<span class="zona-tag">${z}</span>`).join('')||'<span style="color:var(--muted);font-size:11px">—</span>';
    html+=`<tr>
      <td><span class="link-cell" onclick="goC('${esc(p.nombre)}')">${p.nombre}</span></td>
      <td><span class="chip ${hc}">${fH(p.horas)}</span></td>
      <td><span class="chip ${fc}">${fH(p.horas_falt)}</span></td>
      <td style="white-space:nowrap">${dh}</td><td>${zh}</td>
      <td>${buildSvcBtn(p.servicios,rid)}</td>
    </tr>`;
    if(p.servicios&&p.servicios.length){
      html+=`<tr class="svc-row" id="srow-${rid}"><td colspan="6"><div class="svc-inner">
        <table class="svc-table"><thead><tr><th>Local</th><th>Zona</th><th>Días</th><th>Horario</th><th>Hs/día</th><th>Hs/sem</th></tr></thead><tbody>
        ${p.servicios.map(s=>`<tr><td style="font-weight:600">${s.local}</td><td><span class="zona-tag" style="font-size:9px">${s.zona||'—'}</span></td><td>${s.dias.map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join('')}</td><td>${s.entrada} — ${s.salida}</td><td>${fH(s.horas_dia)}</td><td><strong>${fH(s.horas_semana)}</strong></td></tr>`).join('')}
        </tbody></table></div></td></tr>`;
    }
  });
  document.getElementById('bm44').innerHTML=html;
  const totH=d.reduce((a,p)=>a+p.horas,0), totF=d.reduce((a,p)=>a+p.horas_falt,0);
  document.getElementById('tfoot-m44').innerHTML=`<td><strong>TOTALES — ${d.length} personas</strong></td><td><strong>${fH(totH/d.length)} prom.</strong></td><td><strong>${fH(totF)} total faltante</strong></td><td colspan="3"></td>`;
}

// ── MAS 44 ──
function rMas44(){
  const d=D.mas44_enriched;
  document.getElementById('mas44-cnt').textContent=`${d.length} personas`;
  let html='';
  d.forEach((p,i)=>{
    const rid='mas44-'+i;
    const dh=p.dias.map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join('');
    const zh=p.zonas.map(z=>`<span class="zona-tag">${z}</span>`).join('')||'—';
    html+=`<tr>
      <td style="font-size:12.5px">${p.nombre}</td>
      <td><span class="chip chip-green">${fH(p.horas)}</span></td>
      <td><span class="chip chip-amber">+${fH(p.horas_extra)}</span></td>
      <td>${dh}</td><td>${zh}</td><td>${buildSvcBtn(p.servicios,rid)}</td>
    </tr>`;
    if(p.servicios&&p.servicios.length){
      html+=`<tr class="svc-row" id="srow-${rid}"><td colspan="6"><div class="svc-inner">
        <table class="svc-table"><thead><tr><th>Local</th><th>Zona</th><th>Días</th><th>Horario</th><th>Hs/día</th><th>Hs/sem</th></tr></thead><tbody>
        ${p.servicios.map(s=>`<tr><td style="font-weight:600">${s.local}</td><td><span class="zona-tag" style="font-size:9px">${s.zona||'—'}</span></td><td>${s.dias.map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join('')}</td><td>${s.entrada} — ${s.salida}</td><td>${fH(s.horas_dia)}</td><td><strong>${fH(s.horas_semana)}</strong></td></tr>`).join('')}
        </tbody></table></div></td></tr>`;
    }
  });
  document.getElementById('bmas44').innerHTML=html;
  const totH=d.reduce((a,p)=>a+p.horas,0),totE=d.reduce((a,p)=>a+p.horas_extra,0);
  document.getElementById('tfoot-mas44').innerHTML=`<td><strong>TOTALES — ${d.length} personas</strong></td><td><strong>${fH(totH/d.length)} prom.</strong></td><td><strong>+${fH(totE)} total extra</strong></td><td colspan="3"></td>`;
}

// ── EXCLUIDOS ──
function rExcl(){
  const MOT={Admin:'GRUPO ONESTA ADM.',Vidriero:'Vidriero / Limpia Vidrio',Maldonado:'Dpto. Maldonado'};
  const CLS={Admin:'chip-purple',Vidriero:'chip-gray',Maldonado:'chip-amber'};
  document.getElementById('excl-cnt').textContent=D.excluidos.length+' personas excluidas';
  document.getElementById('bexcl').innerHTML=D.excluidos.map(r=>
    `<tr><td style="font-size:12px">${r.Nombre}</td><td><span class="chip ${CLS[r.cat_excl]||'chip-gray'}">${r.cat_excl}</span></td><td>${fH(r.horas)}</td><td style="font-size:11px;color:var(--muted)">${r.locales}</td><td style="font-size:11px;color:var(--muted)">${MOT[r.cat_excl]||'—'}</td></tr>`
  ).join('');
}

// ── MAPS ──
let _mapsInit={svc:false,pers:false};
function getColor(val,max,type){
  const t=max>0?val/max:0;
  if(type==='svc'){
    if(t<.15)return'#ddd6fe';if(t<.3)return'#a78bfa';if(t<.5)return'#7c3aed';if(t<.75)return'#6d28d9';return'#4c1d95';
  } else {
    if(t<.15)return'#d1fae5';if(t<.3)return'#6ee7b7';if(t<.5)return'#10b981';if(t<.75)return'#059669';return'#064e3b';
  }
}
function initMaps(){
  initOneMap('svc');
  initOneMap('pers');
}
function initOneMap(type){
  const elId='map-'+type;
  const el=document.getElementById(elId);
  if(!el){console.warn('Map element not found:',elId);return;}
  if(_mapsInit[type]){
    setTimeout(()=>{if(_maps[type])_maps[type].invalidateSize();},100);
    return;
  }
  const zdata=type==='svc'?D.heatmap_svc_zona_coords:D.heatmap_pers_zona_coords;
  if(!zdata){console.warn('No heatmap data for',type);return;}
  const entries=Object.entries(zdata).filter(([,v])=>v&&v.coords&&v.n>0);
  if(!entries.length){el.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--muted);font-size:13px">Sin datos de coordenadas disponibles</div>';return;}
  const vals=entries.map(([,v])=>v.n);
  const maxVal=Math.max(...vals,1);
  try{
    const map=L.map(elId,{zoomControl:true,scrollWheelZoom:false});
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{
      attribution:'© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
      maxZoom:18
    }).addTo(map);
    _maps[type]=map;
    const bounds=[];
    entries.forEach(([zona,v])=>{
      const r=Math.max(16,Math.min(50,8+v.n*1.8));
      const col=getColor(v.n,maxVal,type);
      const lbl=type==='svc'?`${v.n} servicio${v.n!==1?'s':''}`:
                               `${v.n} persona${v.n!==1?'s':''}`;
      const marker=L.circleMarker([v.coords.lat,v.coords.lon],{
        radius:r,fillColor:col,color:'rgba(255,255,255,0.8)',
        weight:2,opacity:1,fillOpacity:0.82
      }).addTo(map);
      marker.bindPopup(`<div style="font-family:Inter,sans-serif;font-size:13px"><strong>${zona}</strong><br>${lbl}</div>`);
      marker.bindTooltip(`${v.n}`,{permanent:true,direction:'center',className:'map-tooltip'});
      bounds.push([v.coords.lat,v.coords.lon]);
    });
    if(bounds.length)map.fitBounds(bounds,{padding:[30,30]});
    _mapsInit[type]=true;
    const steps=[0.15,0.35,0.6,0.85,1];
    const labSvc=['1–4','5–10','11–20','21–40','40+'];
    const labPers=['1–3','4–8','9–15','16–30','30+'];
    const labs=type==='svc'?labSvc:labPers;
    document.getElementById('legend-'+type).innerHTML=
      '<span style="font-weight:600;margin-right:8px">Intensidad:</span>'+
      steps.map((s,i)=>`<span class="legend-item"><span class="legend-dot" style="background:${getColor(s,1,type)};border-radius:50%"></span>${labs[i]}</span>`).join('')+
      `<span style="margin-left:auto;color:var(--muted);font-size:11px">Tamaño del círculo = cantidad de ${type==='svc'?'servicios':'personas'}</span>`;
  }catch(e){console.error('Map init error:',e);el.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--muted)">Error al cargar el mapa</div>';}
}

// ── ANÁLISIS COMERCIAL ──
function rComercial(){
  const sorted=Object.entries(D.comercial_detail).sort((a,b)=>b[1].n-a[1].n);
  document.getElementById('com-content').innerHTML=sorted.map(([zona,v],zi)=>{
    const n=v.n, ft=Math.round(v.hf_total), fa=Math.round(v.hf_avg);
    const persRows=(v.personas||[]).map(p=>{
      const dTrab=(p.dt||[]).map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join('');
      const dLibre=(p.dl||[]).length?(p.dl).map(d=>`<span class="dia-tag" style="background:#fef3c7;color:#92400e">${d.substring(0,3)}</span>`).join(''):'<span style="color:var(--muted);font-size:11px">Ninguno</span>';
      const svcsHTML=(p.svcs||[]).length?
        (p.svcs).map(s=>`<div style="font-size:11px;margin-bottom:3px"><strong>${s.l}</strong> <span class="zona-tag" style="font-size:9px">${s.z||'—'}</span> ${(s.dias||[]).map(d=>d.substring(0,3)).join(',')} · ${s.e}–${s.s} <strong style="color:var(--accent)">(${s.hs}h/sem)</strong></div>`).join(''):
        '<span style="color:var(--muted);font-size:11px">—</span>';
      const optsHTML=(p.op&&p.op.length)?
        p.op.map(o=>{
          const diasNuevos=(o.dias||[]).filter(d=>!(p.dt||[]).includes(d));
          const diasAmpl=(o.dias||[]).filter(d=>(p.dt||[]).includes(d));
          const diasStr=(o.dias||[]).join(', ');
          const diasDesc=diasNuevos.length&&diasAmpl.length
            ?`${diasNuevos.join(', ')} (días nuevos) + ${diasAmpl.join(', ')} (después del servicio actual)`
            :diasNuevos.length?diasNuevos.join(', ')+' (días nuevos)'
            :diasAmpl.join(', ')+' (después del servicio actual)';
          const franjasHTML=(o.fr||[]).map(f=>`<div style="font-size:10px;color:var(--muted);margin-left:8px">↳ ${f.dia}: ${f.de} a ${f.ha}${f.ref?' (después de '+f.ref+')':''}</div>`).join('');
          return `<div style="margin-bottom:6px"><span class="opcion-tag">${o.n} día${o.n!==1?'s':''}/sem × ${o.h}h = ${o.tot}h</span><div style="font-size:10px;color:var(--muted);margin-top:2px">Días: ${diasStr}</div>${franjasHTML}</div>`;
        }).join(''):
        '<span style="color:var(--muted);font-size:11px">Sin opciones calculadas</span>';
      const nm=esc(p.nombre);
      return `<tr>
        <td style="font-weight:600;font-size:12px"><span class="link-cell" onclick="goC('${nm}')">${p.nombre}</span></td>
        <td><span class="chip chip-red">${fH(p.hf)}</span><br><span style="font-size:10px;color:var(--muted)">${fH(p.ha)} actuales</span></td>
        <td style="white-space:nowrap">${dTrab}</td>
        <td>${dLibre}</td>
        <td style="font-size:11px;line-height:1.6;min-width:200px">${svcsHTML}</td>
        <td style="min-width:220px">${optsHTML}</td>
      </tr>`;
    }).join('');
    return `<div class="com-zone-card">
      <div class="com-zone-header" onclick="toggleZone('zone-${zi}')">
        <div>
          <div class="com-zone-name">${zona}</div>
          <div style="font-size:12px;color:var(--muted);margin-top:2px">${n} persona${n!==1?'s':''} con horas incompletas · ${fa}h faltantes promedio · ${ft}h totales a cubrir</div>
        </div>
        <div style="display:flex;gap:8px;align-items:center">
          <span class="chip chip-red">${ft}h sin cubrir</span>
          <span style="color:var(--muted);font-size:18px" id="arr-${zi}">▶</span>
        </div>
      </div>
      <div class="com-zone-body" id="zone-${zi}">
        <div class="com-zone-summary">
          En <strong>${zona}</strong>, ${n} persona${n!==1?'s tienen':' tiene'} rutina incompleta con promedio de <strong>${fa}h faltantes</strong> — <strong>${ft}h semanales sin cubrir</strong> en total. La columna <em>"Servicio que necesita"</em> muestra exactamente qué días, cuántas horas y en qué franja horaria está disponible cada persona.
        </div>
        <div style="overflow-x:auto">
          <table class="com-persona-table">
            <thead><tr>
              <th>Persona</th><th>Hs faltantes</th><th>Días que trabaja</th>
              <th>Días libres</th><th>Rutina actual completa</th>
              <th>Servicio que necesita (días, horas y franja disponible)</th>
            </tr></thead>
            <tbody>${persRows}</tbody>
          </table>
        </div>
      </div>
    </div>`;
  }).join('');
}
function toggleZone(id){
  const el=document.getElementById(id),zi=id.replace('zone-',''),arr=document.getElementById('arr-'+zi);
  const o=el.classList.toggle('open');arr.textContent=o?'▼':'▶';
}

function toggleZone(id){
  const el=document.getElementById(id),zi=id.replace('zone-',''),arr=document.getElementById('arr-'+zi);
  const o=el.classList.toggle('open');arr.textContent=o?'▼':'▶';
}
// ── CALENDAR MULTI-SELECT ──
let _n='',_hb=0,_sb={},_sw={},_ip=[],_ss=[],_ai=-1,_pf='';
function populateSel(){
  const sel=document.getElementById('csel');
  D.menos44_enriched.forEach(p=>{
    const o=document.createElement('option');o.value=p.nombre;o.text=p.nombre+' ('+fH(p.horas)+')';sel.appendChild(o);
  });
}
function renderCal(nombre){
  if(!nombre){document.getElementById('cal-main').style.display='none';return}
  document.getElementById('cal-main').style.display='block';
  _n=nombre;const ho=D.menos44_enriched.find(p=>p.nombre===nombre);
  _hb=ho?ho.horas:0;
  _sb=JSON.parse(JSON.stringify(SCHED[nombre]||{}));
  _sw=JSON.parse(JSON.stringify(_sb));
  _ip=D.persona_proposals[nombre]||[];
  _ss=[];_ai=-1;_pf='';
  document.getElementById('psearch').value='';
  document.getElementById('pdetail').innerHTML='';
  document.getElementById('cn').textContent=nombre;
  updAll();
}
function resetSel(){_ss=[];_sw=JSON.parse(JSON.stringify(_sb));_ai=-1;document.getElementById('pdetail').innerHTML='';updAll();}
function jsCk(svc){
  if(_ss.some(x=>x.id===svc.id))return false;
  const vs=svc.entrada_h,ve=svc.salida_h;
  if(vs==null||ve==null)return false;
  for(const d of svc.dias){for(const s of (_sw[d]||[])){if(s.entrada_h==null||s.salida_h==null)continue;if(!(ve<=s.entrada_h||vs>=s.salida_h))return false;}}
  const dr=new Set([...Object.keys(_sw),...svc.dias]);
  return DIAS.filter(d=>!dr.has(d)).length>=1;
}
function getCk(){return _ip.filter(p=>{const s=SVC_MAP[p.id];return s&&jsCk(s);})}
function getHS(){return _ss.reduce((a,p)=>{const s=SVC_MAP[p.id];return a+(s?s.horas_semana:0);},0);}
function updAll(){
  const hs=getHS(),ht=_hb+hs,cp=getCk();
  document.getElementById('ch').textContent=`${fH(_hb)} actuales · ${_ss.length} seleccionado${_ss.length!==1?'s':''} · ${cp.length} disponible${cp.length!==1?'s':''}`;
  document.getElementById('cdec').innerHTML=hs===0?'<span class="chip chip-gray">Sin selección</span>':ht<=44?`<span class="chip chip-green">✔ ${fH(ht)}</span>`:`<span class="chip chip-amber">⚠ ${fH(ht)}</span>`;
  // Progress
  const mx=Math.max(ht*1.05,46),bp=Math.min(100,(_hb/mx)*100),ap=Math.min(100-bp,(hs/mx)*100),m44=Math.min(100,(44/mx)*100);
  const bc=_hb<30?'#dc2626':_hb<40?'#d97706':'#2563eb',ac=ht<=44?'#059669':'#d97706';
  document.getElementById('hprog').innerHTML=`<div class="hp-labels"><span>${fH(_hb)} actuales</span>${hs>0?`<span style="color:var(--green)">+${fH(hs)} seleccionados = <strong>${fH(ht)}</strong></span>`:''}<span>objetivo 44h</span></div><div class="hp-track"><div class="hp-base" style="width:${bp}%;background:${bc}"></div><div class="hp-added" style="left:${bp}%;width:${ap}%;background:${ac}"></div><div class="hp-mark" style="left:${m44}%"></div></div>`;
  // Pills
  const dt=Object.keys(_sw),dl=DIAS.filter(d=>!dt.includes(d));
  const hsExtra=ht>44?Math.round((ht-44)*10)/10:0;
  document.getElementById('cpills').innerHTML=[
    {l:'Horas base',v:fH(_hb),c:_hb<30?'v-red':_hb<40?'v-amber':''},
    {l:'Hs seleccionadas',v:hs>0?'+'+fH(hs):'—',c:hs>0?'v-green':''},
    {l:'Total proyectado',v:fH(ht),c:hs>0?(ht<=44?'v-green':'v-amber'):''},
    {l:'Faltante',v:ht>=44?'✔ Completo':fH(44-ht),c:ht>=44?'v-green':''},
    {l:'Horas extra/sem',v:hsExtra>0?'+'+fH(hsExtra):'—',c:hsExtra>0?'v-amber':''},
    {l:'Días trabajados',v:dt.length+'/7',c:''},
    {l:'Días libres',v:dl.length?dl.map(d=>d.substring(0,3)).join(', '):'—',c:''},
  ].map(i=>`<div class="pill"><div class="pill-label">${i.l}</div><div class="pill-value ${i.c}">${i.v}</div></div>`).join('');
  // Strip
  document.getElementById('btn-rst').style.display=_ss.length?'':'none';
  document.getElementById('strip-tot').innerHTML=`${fH(ht)} <span>totales</span>`;
  document.getElementById('strip-sum').textContent=_ss.length?`${_ss.length} servicio${_ss.length!==1?'s':''} · +${fH(hs)}`:'';
  document.getElementById('strip-items').innerHTML=_ss.length?_ss.map((p,i)=>{const s=SVC_MAP[p.id];if(!s)return'';return`<div class="strip-item-card"><div><div class="sic-local">${s.local}</div><div class="sic-meta">${s.dias.join(', ')} · ${s.entrada}–${s.salida} · ${fH(s.horas_semana)}</div></div><button class="remove-btn" onclick="remSel(${i})">✕</button></div>`;}).join(''):'<div style="font-size:12px;color:var(--muted);font-style:italic">Ningún servicio agregado — hacé clic en "+ Agregar"</div>';
  rPropList(cp);rGrid();
  const rs=document.getElementById('route-section');
  if(rs&&_n){rs.style.display='block';setTimeout(()=>{renderRouteMap();renderBoletoTable();},80);}
}
function remSel(i){const r=_ss.splice(i,1)[0];const s=SVC_MAP[r.svc_id];if(s)s.dias.forEach(d=>{if(_sw[d]){_sw[d]=_sw[d].filter(x=>!(x.local===s.local&&x.entrada===s.entrada));if(!_sw[d].length)delete _sw[d];}});updAll();}
function addS(i){
  const cp=getCk(),q=_pf.toLowerCase();
  const fl=q?cp.filter(p=>{const s=SVC_MAP[p.id];return s&&(s.local.toLowerCase().includes(q)||s.zona.toLowerCase().includes(q));}):cp;
  const p=fl[i];if(!p)return;const s=SVC_MAP[p.id];if(!s)return;
  _ss.push(p);s.dias.forEach(d=>{if(!_sw[d])_sw[d]=[];_sw[d].push({local:s.local,direccion:s.direccion,entrada:s.entrada,salida:s.salida,entrada_h:s.entrada_h,salida_h:s.salida_h,horas:s.horas_dia});});
  _ai=-1;document.getElementById('pdetail').innerHTML='';updAll();
}
function rPropList(cp){
  document.getElementById('plbl').textContent=`Servicios compatibles disponibles (${cp.length})`;
  const q=_pf.toLowerCase();
  const fl=q?cp.filter(p=>{const s=SVC_MAP[p.id];return s&&(s.local.toLowerCase().includes(q)||s.zona.toLowerCase().includes(q)||(s.direccion||'').toLowerCase().includes(q));}):cp;
  if(!fl.length){document.getElementById('plist').innerHTML=`<div class="prop-empty">${cp.length?'Sin resultados':'No hay servicios compatibles'}</div>`;return;}
  document.getElementById('plist').innerHTML=fl.map((p,i)=>{
    const s=SVC_MAP[p.id];if(!s)return'';
    const ht=_hb+getHS()+s.horas_semana,isR=ht>44,isO=s.solo_sobre44,isA=_ai===i;
    return`<div class="prop-card${isO?' opt':''}${isA?' active':''}" onclick="showD(${i})">
      ${isO?'<div class="pc-opt">⭐ Optimiza distribución</div>':''}
      <div class="pc-local">${s.local}</div>
      <div class="pc-dias">📅 ${s.dias.join(', ')} · ${s.n_dias} día${s.n_dias!==1?'s':''}</div>
      <div class="pc-meta">${s.entrada}–${s.salida}${s.zona?' · '+s.zona:''}${p.km?' · '+p.km+'km':''}</div>
      <div class="pc-hours">+${fH(s.horas_semana)} → ${fH(ht)}</div>
      <div class="pc-badges"><span class="pc-badge ${isR?'pbr':'pbs'}">${isR?'⚠ REVISAR':'✔ SUGERIDO'}</span><span class="pc-badge ${p.bus===false?'pbw':'pbb'}">${p.bus===false?'🚶':'🚌'}</span></div>
      <button class="add-btn-sm" onclick="event.stopPropagation();addS(${i})">+ Agregar</button>
    </div>`;
  }).join('');
}
function filterP(q){_pf=q;rPropList(getCk());}
function showD(i){
  _ai=i;const cp=getCk(),q=_pf.toLowerCase();
  const fl=q?cp.filter(p=>{const s=SVC_MAP[p.id];return s&&(s.local.toLowerCase().includes(q)||s.zona.toLowerCase().includes(q));}):cp;
  const p=fl[i];if(!p)return;const s=SVC_MAP[p.id];if(!s)return;
  const ht=_hb+getHS()+s.horas_semana,isR=ht>44;
  const dt=p.km?p.km+' km (~'+p.c+' cuadras)':'Sin coordenadas';
  document.getElementById('pdetail').innerHTML=`<div class="det-box ${isR?'det-rev':'det-sug'}">
    ${isR?'<strong>⚠ Superaría las 44h.</strong><br>':''}
    ${s.solo_sobre44?'<strong>⭐ Optimización:</strong> actualmente solo cubierto por personal +44h.<br>':''}
    ${s.personas_actuales&&s.personas_actuales.length?`Cubierto por: <em>${s.personas_actuales.join(', ')}</em><br>`:''}
    <strong>${s.local}</strong> · ${s.dias.join(', ')} · ${s.entrada}–${s.salida} · Suma <strong>${fH(s.horas_semana)}</strong> → total <strong>${fH(ht)}</strong> · ${(p.bus===null?'Revisar':p.bus?'Revisar / posible ómnibus':'OK caminable')} · Boleto: ${(p.bus===null?'Revisar':p.bus?'Sí':'No')}
    <div class="det-grid">${[['Local',s.local],['Zona',s.zona||'—'],['Dirección',s.direccion],['Días',s.dias.join(', ')],['Días/sem',s.n_dias],['Horario',s.entrada+' — '+s.salida],['Hs/día',fH(s.horas_dia)],['Hs/semana',fH(s.horas_semana)],['Total proy.',fH(ht)],['Distancia',dt],['Traslado',(p.bus===null?'Revisar':p.bus?'Revisar / posible ómnibus':'OK caminable')],['Boleto',(p.bus===null?'Revisar':p.bus?'Sí':'No')]].map(([l,v])=>`<div class="det-item"><div class="det-lbl">${l}</div><div class="det-val">${v}</div></div>`).join('')}</div>
    <button onclick="addS(${i})" style="margin-top:10px;background:var(--accent);color:#fff;border:none;border-radius:var(--radius-sm);padding:8px 20px;font-size:12px;font-weight:700;cursor:pointer;font-family:'Inter',sans-serif">+ Agregar este servicio</button>
  </div>`;
  rPropList(cp);rGrid();
}
function rGrid(){
  const sd={};_ss.forEach(p=>{const s=SVC_MAP[p.id];if(s)s.dias.forEach(d=>{if(!sd[d])sd[d]=[];sd[d].push(s);});});
  const cp=getCk(),q=_pf.toLowerCase();
  const fl=q?cp.filter(p=>{const s=SVC_MAP[p.id];return s&&(s.local.toLowerCase().includes(q)||s.zona.toLowerCase().includes(q));}):cp;
  const prev=_ai>=0?SVC_MAP[(fl[_ai]||{}).svc_id]:null,pd=prev?prev.dias:[];
  document.getElementById('cgrid').innerHTML=DIAS.map((dia,di)=>{
    const base=_sb[dia]||[],hs=!!sd[dia],hn=pd.indexOf(dia)!==-1;
    let bl='';
    base.forEach(s=>{bl+=`<div class="cg-block cb-act"><strong>${s.entrada}–${s.salida}</strong><div style="font-size:10px;color:var(--text2)">${s.local}</div><div class="cg-type" style="color:var(--blue)">● Actual · ${fH(s.horas)}</div></div>`;});
    if(hs)sd[dia].forEach(s=>{bl+=`<div class="cg-block cb-sel"><strong>${s.entrada}–${s.salida}</strong><div style="font-size:10px;color:var(--text2)">${s.local}</div><div class="cg-type" style="color:var(--accent)">✔ Seleccionado · ${fH(s.horas_dia)}</div></div>`;});
    if(hn&&prev)bl+=`<div class="cg-block cb-new"><strong>${prev.entrada}–${prev.salida}</strong><div style="font-size:10px;color:var(--text2)">${prev.local}</div><div class="cg-type" style="color:var(--green)">👁 Vista previa · ${fH(prev.horas_dia)}</div></div>`;
    const dc=hn?'dp-new':hs?'dp-sel':'';
    return`<div class="cg-day ${dc}"><div class="cg-dn">${DIAS3[di]}</div>${!base.length&&!hs&&!hn?`<div class="cg-libre"><span>Libre</span></div>`:bl}</div>`;
  }).join('');
}
function goC(n){
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-calendario').classList.add('active');
  document.querySelectorAll('.nav-item')[4].classList.add('active');
  document.getElementById('csel').value=n;renderCal(n);
}

// ── PERSONAS LIST ──
let _ap=[];
function rPers(){_ap=D.personas_simple;drawP(_ap);}
function drawP(rows){
  document.getElementById('bpers').innerHTML=rows.map(r=>{
    const im=r.horas_falt<=0;
    const hc=im?'chip-green':r.horas_act<30?'chip-red':r.horas_act<40?'chip-amber':'chip-blue';
    const fc=im?'chip-green':r.horas_falt>10?'chip-red':r.horas_falt>5?'chip-amber':'chip-green';
    const dh=r.dias&&r.dias.length?r.dias.map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join(''):'—';
    const zh=r.zonas&&r.zonas.length?r.zonas.map(z=>`<span class="zona-tag">${z}</span>`).join(''):'—';
    const nc=r.dias&&r.dias.length&&!im?`<span class="link-cell" onclick="goC('${esc(r.nombre)}')">${r.nombre}</span>`:r.nombre;
    return`<tr><td>${nc}</td><td><span class="chip ${hc}">${fH(r.horas_act)}</span></td><td>${im?'<span class="chip chip-green">—</span>':`<span class="chip ${fc}">${fH(r.horas_falt)}</span>`}</td><td>${dh}</td><td>${zh}</td></tr>`;
  }).join('');
  const tot=rows.length, totH=rows.reduce((a,r)=>a+r.horas_act,0);
  document.getElementById('tfoot-pers').innerHTML=`<td><strong>Total: ${tot} personas</strong></td><td><strong>${fH(totH/Math.max(1,tot))} prom.</strong></td><td></td><td colspan="3"></td>`;
}
function fPers(q){q=q.toLowerCase();drawP(_ap.filter(r=>r.nombre.toLowerCase().includes(q)))}
function fPersZona(q){q=q.toLowerCase();drawP(_ap.filter(r=>!q||(r.zonas&&r.zonas.some(z=>z.toLowerCase().includes(q)))))}
function fPersGrp(v){if(!v){drawP(_ap);return}drawP(_ap.filter(r=>v==='menos44'?r.horas_falt>0:r.horas_falt<=0))}

// ── TABLE UTILS ──
function fTbl(id,q){q=q.toLowerCase();document.querySelectorAll('#'+id+' tbody tr:not(.svc-row)').forEach(r=>r.style.display=r.textContent.toLowerCase().includes(q)?'':'none')}
function sTbl(id,col){
  const tb=document.querySelector('#'+id+' tbody'),rows=[...tb.querySelectorAll('tr:not(.svc-row)')];
  const dir=tb.dataset.sd==='asc'?-1:1;tb.dataset.sd=dir===1?'asc':'desc';
  rows.sort((a,b)=>{
    const av=(a.cells[col]?.textContent||'').trim().replace(/[^0-9.]/g,'');
    const bv=(b.cells[col]?.textContent||'').trim().replace(/[^0-9.]/g,'');
    const an=parseFloat(av),bn=parseFloat(bv);
    if(!isNaN(an)&&!isNaN(bn))return dir*(an-bn);
    return dir*(a.cells[col]?.textContent||'').localeCompare(b.cells[col]?.textContent||'','es');
  });
  rows.forEach(r=>tb.appendChild(r));
}

// ── DISPONIBILIDAD POR HORA ──
function initDispSelects(){
  // Hora select: 06:00 to 21:00
  const hSel=document.getElementById('disp-hora');
  for(let h=6;h<=21;h++){
    const o=document.createElement('option');
    o.value=h;o.text=`${String(h).padStart(2,'0')}:00`;
    hSel.appendChild(o);
  }
  hSel.value=9; // default 09:00
  
  // Zona select
  const zSel=document.getElementById('disp-zona');
  const zonas=new Set();
  Object.values(D.persona_disponibilidad||{}).forEach(p=>p.zonas.forEach(z=>zonas.add(z)));
  [...zonas].sort().forEach(z=>{
    const o=document.createElement('option');o.value=z;o.text=z;zSel.appendChild(o);
  });
  
  // Trigger on dia/dur change
  document.getElementById('disp-dia').onchange=queryDisp;
  document.getElementById('disp-dur').onchange=queryDisp;
  document.getElementById('disp-zona').onchange=queryDisp;
}

function queryDisp(){
  const dia=document.getElementById('disp-dia').value;
  const hora=parseFloat(document.getElementById('disp-hora').value);
  const zona=document.getElementById('disp-zona').value;
  const pd=PDISP;
  const results=[];
  Object.entries(pd).forEach(([nombre,info])=>{
    if(zona&&!(info.zonas||[]).includes(zona))return;
    const dayAvail=(info.disponibilidad||{})[dia]||[];
    for(const slot of dayAvail){
      const desde=parseTime(slot.desde);
      const hasta=parseTime(slot.hasta);
      if(desde<=hora&&hasta>hora){
        const dispH=Math.round((hasta-hora)*10)/10;
        results.push({nombre,horas_falt:info.horas_falt,zonas:info.zonas||[],
          libre_desde:slot.desde,libre_hasta:slot.hasta,horas_disponibles:dispH});
        break;
      }
    }
  });
  results.sort((a,b)=>b.horas_disponibles-a.horas_disponibles);
  const hStr=(hora<10?'0':'')+hora+':00';
  document.getElementById('disp-result-header').style.display='block';
  document.getElementById('disp-title').textContent=results.length+' persona'+(results.length!==1?'s':'')+' disponible'+(results.length!==1?'s':'')+' el '+dia+' a las '+hStr+(zona?' en '+zona:'');
  document.getElementById('disp-sub').textContent='Disponibles a las '+hStr+(zona?' · Zona: '+zona:' · Todas las zonas')+' · Ordenado por más horas libres ese día';
  if(!results.length){
    document.getElementById('disp-panel').style.display='none';
    document.getElementById('disp-empty').style.display='block';
    return;
  }
  document.getElementById('disp-panel').style.display='block';
  document.getElementById('disp-empty').style.display='none';
  const DIAS_ORD=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
  document.getElementById('tbody-disp').innerHTML=results.map(r=>{
    const sched=SCHED[r.nombre]||{};
    const diasTrab=Object.keys(sched).sort((a,b)=>DIAS_ORD.indexOf(a)-DIAS_ORD.indexOf(b));
    const diasH=diasTrab.map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join('');
    const zonasH=r.zonas.length?r.zonas.map(z=>`<span class="zona-tag">${z}</span>`).join(''):'<span style="color:var(--muted);font-size:11px">—</span>';
    const faltC=r.horas_falt>10?'chip-red':r.horas_falt>5?'chip-amber':'chip-green';
    const nm=esc(r.nombre);
    return `<tr>
      <td style="font-weight:600;font-size:12.5px"><span class="link-cell" onclick="goC('${nm}')">${r.nombre}</span></td>
      <td><span class="chip ${faltC}">${fH(r.horas_falt)}</span></td>
      <td style="font-weight:700;color:var(--green);font-size:13px">${r.libre_desde}</td>
      <td style="color:var(--muted)">${r.libre_hasta}</td>
      <td><span class="chip chip-blue">${fH(r.horas_disponibles)}</span></td>
      <td>${diasH}</td>
      <td>${zonasH}</td>
    </tr>`;
  }).join('');
}


function initDispSelects(){
  const hSel=document.getElementById('disp-hora');
  for(let h=6;h<=21;h++){
    const o=document.createElement('option');
    o.value=h;
    o.text=(h<10?'0':'')+h+':00';
    hSel.appendChild(o);
  }
  hSel.value='9';
  const zSel=document.getElementById('disp-zona');
  const zonas=new Set();
  Object.values(D.persona_disponibilidad||{}).forEach(p=>(p.zonas||[]).forEach(z=>{if(z)zonas.add(z)}));
  [...zonas].sort().forEach(z=>{const o=document.createElement('option');o.value=z;o.text=z;zSel.appendChild(o);});
  document.getElementById('disp-dia').onchange=queryDisp;
  document.getElementById('disp-dur').onchange=queryDisp;
  document.getElementById('disp-zona').onchange=queryDisp;
}

function parseTime(s){
  if(!s)return 0;
  const p=s.split(':');
  return parseInt(p[0])+(parseInt(p[1]||0)/60);
}


function haversineDist(lat1,lon1,lat2,lon2){
  const R=6371000,dL=(lat2-lat1)*Math.PI/180,dO=(lon2-lon1)*Math.PI/180;
  const a=Math.sin(dL/2)**2+Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dO/2)**2;
  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
function getCoords(c){
  if(!c)return null;
  if(typeof c.lat!=='undefined')return[c.lat,c.lon];
  if(Array.isArray(c)&&c.length===2)return[c[0],c[1]];
  return null;
}

function renderRouteMap(){
  const el=document.getElementById('route-map');
  if(!el)return;
  const seen=new Set(),allSvcs=[];
  // base services
  Object.values(_sb).flat().forEach(s=>{
    const coords=s.coords||LOCAL_COORDS[(s.local||'').trim().toUpperCase()];
    if(!seen.has(s.local)&&coords){const c=getCoords(coords);if(c){seen.add(s.local);allSvcs.push({local:s.local,c,tipo:'actual',col:'#2563eb'});}}
  });
  // selected
  _ss.forEach(p=>{
    const s=SVC_MAP[p.id];
    if(s&&s.coords&&!seen.has(s.local)){const c=getCoords(s.coords);if(c){seen.add(s.local);allSvcs.push({local:s.local,c,tipo:'sel',col:'#059669'});}}
  });
  if(!allSvcs.length){el.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--muted);font-size:12px">Sin coordenadas disponibles</div>';return;}
  if(_routeMap){try{_routeMap.remove();}catch(e){}}_routeMap=null;
  el.innerHTML='';
  try{
    const map=L.map('route-map',{zoomControl:true,scrollWheelZoom:false});
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap',maxZoom:18}).addTo(map);
    _routeMap=map;
    allSvcs.forEach((s,i)=>{
      L.circleMarker(s.c,{radius:11,fillColor:s.col,color:'#fff',weight:2,fillOpacity:.9})
        .addTo(map).bindPopup('<strong>'+s.local+'</strong>');
      L.tooltip({permanent:true,direction:'center',className:'map-tooltip'})
        .setContent('<span style="font-size:10px;font-weight:800;color:#fff">'+(i+1)+'</span>')
        .setLatLng(s.c).addTo(map);
    });
    if(allSvcs.length>1)
      L.polyline(allSvcs.map(s=>s.c),{color:'#6366f1',weight:2,opacity:.5,dashArray:'5,4'}).addTo(map);
    map.fitBounds(allSvcs.map(s=>s.c),{padding:[28,28]});
  }catch(e){console.error(e);}
}

function renderBoletoTable(){
  const el=document.getElementById('boleto-table-wrap');
  if(!el)return;
  const DIAS_ORD=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
  const seq=SCHED[_n]||{};
  const selByDay={};
  _ss.forEach(p=>{
    const s=SVC_MAP[p.id];
    if(!s)return;
    s.dias.forEach(dia=>{
      if(!selByDay[dia])selByDay[dia]=[];
      const c=s.coords?getCoords(s.coords):null;
      selByDay[dia].push({local:s.local,entrada:s.entrada,salida:s.salida,coords:c,tipo:'sel'});
    });
  });
  const rows=[];
  DIAS_ORD.forEach(dia=>{
    const base=(seq[dia]||[]).map(s=>({...s,tipo:'act',coords:s.coords?getCoords(s.coords):(LOCAL_COORDS[(s.local||'').trim().toUpperCase()]?getCoords(LOCAL_COORDS[(s.local||'').trim().toUpperCase()]):null)}));
    const sel=selByDay[dia]||[];
    const all=[...base,...sel].sort((a,b)=>{
      const ta=a.entrada?+a.entrada.replace(':',''):0;
      const tb=b.entrada?+b.entrada.replace(':',''):0;
      return ta-tb;
    });
    for(let i=1;i<all.length;i++){
      const p=all[i-1],c=all[i];
      let cuadras=null;
      if(p.coords&&c.coords){
        const d=haversineDist(p.coords[0],p.coords[1],c.coords[0],c.coords[1]);
        cuadras=Math.round(d/100);
      }
      rows.push({dia,de:p.local,de_sal:p.salida,a:c.local,a_ent:c.entrada,cuadras,boleto:cuadras===null?null:cuadras>10});
    }
  });
  if(!rows.length){el.innerHTML='<div style="padding:10px 0;color:var(--muted);font-size:12px">Sin traslados entre servicios o sin coordenadas.</div>';return;}
  const nBoleto=rows.filter(r=>r.boleto).length;
  el.innerHTML=
    '<div style="margin-bottom:10px;font-size:12px">'
    +(nBoleto?'<span style="background:#fef3c7;color:#92400e;padding:3px 10px;border-radius:6px;font-weight:700;margin-right:8px">🎫 '+nBoleto+' traslado'+(nBoleto!==1?'s':'')+' con boleto</span>':'')
    +'<span style="background:#d1fae5;color:#065f46;padding:3px 10px;border-radius:6px;font-weight:600">'+(rows.filter(r=>r.boleto===false).length)+' caminable'+(rows.filter(r=>r.boleto===false).length!==1?'s':'')+'</span></div>'
    +'<table style="width:100%;border-collapse:collapse;font-size:12px">'
    +'<thead><tr style="background:var(--surface2)"><th style="padding:7px 10px;text-align:left;font-size:10px;color:var(--muted);text-transform:uppercase;border-bottom:1px solid var(--border)">Día</th><th style="padding:7px 10px;text-align:left;font-size:10px;color:var(--muted);text-transform:uppercase;border-bottom:1px solid var(--border)">Desde</th><th style="padding:7px 10px;text-align:left;font-size:10px;color:var(--muted);text-transform:uppercase;border-bottom:1px solid var(--border)">Hacia</th><th style="padding:7px 10px;text-align:center;font-size:10px;color:var(--muted);text-transform:uppercase;border-bottom:1px solid var(--border)">Cuadras</th><th style="padding:7px 10px;text-align:center;font-size:10px;color:var(--muted);text-transform:uppercase;border-bottom:1px solid var(--border)">Boleto</th></tr></thead>'
    +'<tbody>'+rows.map(r=>
      '<tr>'
      +'<td style="padding:7px 10px;border-bottom:1px solid var(--border);font-weight:600">'+r.dia.substring(0,3)+'</td>'
      +'<td style="padding:7px 10px;border-bottom:1px solid var(--border)"><div style="font-weight:600;font-size:11.5px">'+r.de+'</div><div style="font-size:10px;color:var(--muted)">Sale '+r.de_sal+'</div></td>'
      +'<td style="padding:7px 10px;border-bottom:1px solid var(--border)"><div style="font-weight:600;font-size:11.5px">'+r.a+'</div><div style="font-size:10px;color:var(--muted)">Entra '+r.a_ent+'</div></td>'
      +'<td style="padding:7px 10px;border-bottom:1px solid var(--border);text-align:center">'+(r.cuadras!==null?r.cuadras+'c':'<span style="color:var(--muted)">—</span>')+'</td>'
      +'<td style="padding:7px 10px;border-bottom:1px solid var(--border);text-align:center">'
        +(r.boleto===null?'<span style="color:var(--muted);font-size:10px">—</span>'
          :r.boleto?'<span style="background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:5px;font-weight:700;font-size:11px">🎫 Sí</span>'
                   :'<span style="background:#d1fae5;color:#065f46;padding:2px 8px;border-radius:5px;font-weight:600;font-size:11px">🚶 No</span>')
        +'</td>'
      +'</tr>'
    ).join('')+'</tbody></table>';
}


// ══════════════════════════════════════════
// BUSCAR POR SERVICIO
// ══════════════════════════════════════════
let _bsMap=null;

function initBuscarServicio(){
  const sel=document.getElementById('bs-local');
  // Unique locals from svc_catalog
  const locals=[...new Set(D.svc_catalog.map(s=>s.l||s.local))].sort();
  locals.forEach(l=>{
    const o=document.createElement('option');
    o.value=l;o.text=l;sel.appendChild(o);
  });
  // Hour select
  const hSel=document.getElementById('bs-hora');
  for(let h=0;h<=23;h++){
    const o=document.createElement('option');
    o.value=h;o.text=(h<10?'0':'')+h+':00';hSel.appendChild(o);
  }
}

function bsOnLocalChange(){
  const local=document.getElementById('bs-local').value;
  if(!local)return;
  const svcs=D.svc_catalog.filter(s=>(s.l||s.local)===local);
  if(!svcs.length)return;
  const s=svcs[0];
  const info=document.getElementById('bs-svc-info');
  info.style.display='block';
  info.innerHTML=`<strong>${local}</strong> · ${s.z||s.zona||'—'} · ${(s.di||s.dias||[]).join(', ')} · ${s.e||s.entrada}–${s.s||s.salida} · ${(s.hs||s.horas_semana)}h/sem · Dirección: ${s.d||s.direccion}`;
}

function buscarPorServicio(){
  const local=document.getElementById('bs-local').value;
  const diaFilt=document.getElementById('bs-dia').value;
  const horaFilt=parseFloat(document.getElementById('bs-hora').value||'-1');
  if(!local){alert('Seleccioná un servicio');return;}

  // Find the service(s) for this local
  const svcs=D.svc_catalog.filter(s=>(s.l||s.local)===local);
  if(!svcs.length)return;
  // If dia filter, pick the svc with that dia
  let targetSvc=svcs[0];
  if(diaFilt){
    const match=svcs.find(s=>(s.di||s.dias||[]).includes(diaFilt));
    if(match)targetSvc=match;
  }

  const svcCoords=targetSvc.c||targetSvc.coords;
  const svcDias=targetSvc.di||targetSvc.dias||[];
  const svcEntrada=parseFloat((targetSvc.e||targetSvc.entrada||'0').split(':')[0]);
  const svcSalida=parseFloat((targetSvc.s||targetSvc.salida||'0').split(':')[0]);
  const DIAS_ORD=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];

  // For each person in menos44: check compatibility & compute distance
  const results=[];
  (D.person_index||[]).forEach(p=>{
    const pCoords=p.c;
    let dist=null,cuadras=null;
    if(svcCoords&&pCoords){
      dist=haversineDist(pCoords[0],pCoords[1],svcCoords[0],svcCoords[1]);
      cuadras=Math.round(dist/100);
    }

    // Check availability on service days
    const diasDisp=[];
    const diasConflicto=[];
    svcDias.forEach(dia=>{
      const daySched=p.ds[dia];
      if(!daySched){
        diasDisp.push({dia,tipo:'libre',desde:'06:00',hasta:'22:00'});
        return;
      }
      // Check if service time fits: before first or after last
      const last=daySched.l||0;
      const first=daySched.f||24;
      if(svcSalida<=first){
        diasDisp.push({dia,tipo:'antes',desde:(targetSvc.e||targetSvc.entrada),hasta:(targetSvc.s||targetSvc.salida)});
      } else if(svcEntrada>=last){
        diasDisp.push({dia,tipo:'despues',desde:(targetSvc.e||targetSvc.entrada),hasta:(targetSvc.s||targetSvc.salida)});
      } else {
        diasConflicto.push(dia);
      }
    });

    const disponible=diasDisp.length===svcDias.length; // all days OK
    const parcial=diasDisp.length>0&&diasConflicto.length>0;

    results.push({
      nombre:p.n, ha:p.ha, hf:p.hf,
      coords:pCoords, dist, cuadras,
      boleto:cuadras===null?null:cuadras>10,
      diasDisp, diasConflicto, disponible, parcial,
      diasLibres:DIAS_ORD.filter(d=>!p.ds[d]),
    });
  });

  // Sort: available first, then by distance
  results.sort((a,b)=>{
    if(a.disponible!==b.disponible)return a.disponible?-1:1;
    if(a.parcial!==b.parcial)return a.parcial?-1:1;
    return (a.dist||9999)-(b.dist||9999);
  });

  // Render table
  document.getElementById('bs-results').style.display='block';
  document.getElementById('bs-result-title').textContent=`${results.filter(r=>r.disponible).length} personas disponibles para "${local}"`;
  document.getElementById('bs-result-sub').textContent=`${results.filter(r=>r.parcial).length} con disponibilidad parcial · ${results.filter(r=>!r.disponible&&!r.parcial).length} con conflicto`;

  document.getElementById('tbody-bs').innerHTML=results.slice(0,50).map(r=>{
    const distTxt=r.dist?Math.round(r.dist)+'m':'—';
    const cuadTxt=r.cuadras!==null?r.cuadras+'c':'—';
    const boletoTxt=r.boleto===null?'—':r.boleto?'<span class="chip ca">🎫 Sí</span>':'<span class="chip cg">🚶 No</span>';
    const hfChip=r.hf>10?'cred':r.hf>5?'ca':'cg';
    const dispTxt=r.disponible
      ?'<span class="chip cg">✔ Todos los días</span>'
      :r.parcial
        ?`<span class="chip ca">⚠ Parcial (${r.diasDisp.map(d=>d.dia.substring(0,3)).join(',')} OK)</span>`
        :'<span class="chip cred">✖ Conflicto</span>';
    const diasLibH=r.diasLibres.length?r.diasLibres.map(d=>`<span class="dia-tag">${d.substring(0,3)}</span>`).join(''):'—';
    const psInfo=D.personas_simple.find(ps=>ps.nombre===r.nombre);
    const zonasH=psInfo&&psInfo.zonas.length?psInfo.zonas.map(z=>`<span class="zona-tag">${z}</span>`).join(''):'—';
    const nm=esc(r.nombre);
    return `<tr>
      <td style="font-size:12px;font-weight:600"><span class="link-cell" onclick="goC('${nm}')">${r.nombre}</span></td>
      <td>${distTxt}</td>
      <td>${cuadTxt}</td>
      <td>${boletoTxt}</td>
      <td><span class="chip ${hfChip}">${fH(r.hf)}</span></td>
      <td>${dispTxt}</td>
      <td>${diasLibH}</td>
      <td>${zonasH}</td>
      <td><button class="btn btn-primary" style="padding:4px 10px;font-size:11px" onclick="goC('${nm}')">Ver calendario</button></td>
    </tr>`;
  }).join('');

  // Render map
  renderBsMap(local, svcCoords, results.slice(0,50));
}

function renderBsMap(local, svcCoords, results){
  const el=document.getElementById('bs-map');
  if(!el)return;
  if(_bsMap){try{_bsMap.remove();}catch(e){}}_bsMap=null;
  el.innerHTML='';
  try{
    const map=L.map('bs-map',{zoomControl:true,scrollWheelZoom:false});
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap',maxZoom:18}).addTo(map);
    _bsMap=map;
    const bounds=[];

    // Service marker (red/star)
    if(svcCoords){
      const sc=Array.isArray(svcCoords)?svcCoords:[svcCoords.lat,svcCoords.lon];
      L.circleMarker(sc,{radius:14,fillColor:'#dc2626',color:'#fff',weight:3,fillOpacity:1})
        .addTo(map).bindPopup('<strong>🎯 '+local+'</strong><br><em>Servicio buscado</em>');
      L.tooltip({permanent:true,direction:'center',className:'map-tooltip'})
        .setContent('<span style="font-size:11px;font-weight:800;color:#fff">★</span>')
        .setLatLng(sc).addTo(map);
      bounds.push(sc);
    }

    // Person markers
    results.forEach((r,i)=>{
      if(!r.coords)return;
      const pc=r.coords;
      const col=r.disponible?'#2563eb':r.parcial?'#d97706':'#6b7280';
      L.circleMarker(pc,{radius:9,fillColor:col,color:'#fff',weight:2,fillOpacity:0.85})
        .addTo(map)
        .bindPopup('<strong>'+r.nombre+'</strong><br>'+(r.cuadras!==null?r.cuadras+'c · ':'')+(r.boleto?'🎫 Boleto':'🚶 Caminable')+'<br><span style="color:'+col+'">'+(r.disponible?'✔ Disponible':r.parcial?'⚠ Parcial':'✖ Conflicto')+'</span>');
      bounds.push(pc);
    });

    if(bounds.length)map.fitBounds(bounds,{padding:[25,25]});
  }catch(e){console.error(e);}
}

// ══════════════════════════════════════════
// GESTOR DE LUGARES
// ══════════════════════════════════════════
let _glMap=null;
let _glMarker=null;
let _glLugares=[...(D.lugares_index||[])]; // working copy (can be edited)

function initGestorLugares(){
  renderGlTable();
  document.getElementById('gl-list-title').textContent='Lugares registrados ('+_glLugares.length+')';
}

function renderGlTable(filter=''){
  const q=filter.toLowerCase();
  const rows=_glLugares.filter(l=>!q||l.nombre.toLowerCase().includes(q)||l.direccion.toLowerCase().includes(q)||(l.zona||'').toLowerCase().includes(q));
  document.getElementById('tbody-gl').innerHTML=rows.map((l,i)=>
    '<tr>'
    +'<td style="font-size:12px;font-weight:600">'+l.nombre+'</td>'
    +'<td style="font-size:11px;color:var(--muted)">'+l.direccion+'</td>'
    +'<td><span class="zona-tag" style="font-size:9px">'+( l.zona||'—')+'</span></td>'
    +'<td style="font-size:11px">'+(l.lat&&l.lon?'<span class="chip cg" style="font-size:9px">✔ '+l.lat.toFixed(4)+'</span>':'<span class="chip cred" style="font-size:9px">Sin coords</span>')+'</td>'
    +'<td><button class="btn" style="padding:3px 10px;font-size:11px;background:var(--surface2);border:1px solid var(--border);border-radius:6px;cursor:pointer" onclick="glEditar('+i+')">✏ Editar</button></td>'
    +'</tr>'
  ).join('');
}

function glFilter(q){renderGlTable(q);}

function glEditar(idx){
  const q=document.querySelector('#tbl-gl tbody input[type=text]')?.value||'';
  const filtered=_glLugares.filter(l=>!q||l.nombre.toLowerCase().includes(q.toLowerCase()));
  const l=filtered[idx];
  if(!l)return;
  document.getElementById('gl-nombre').value=l.nombre||'';
  document.getElementById('gl-dir').value=l.direccion||'';
  document.getElementById('gl-ciudad').value=l.ciudad||'Montevideo';
  document.getElementById('gl-zona').value=l.zona||'';
  document.getElementById('gl-lat').value=l.lat||'';
  document.getElementById('gl-lon').value=l.lon||'';
  if(l.lat&&l.lon)glShowOnMap(l.lat,l.lon,l.nombre);
  glStatus('Lugar cargado para edición. Corregí lo que necesites y guardá.','info');
}

async function glGeocodificar(){
  const dir=document.getElementById('gl-dir').value.trim();
  const ciudad=document.getElementById('gl-ciudad').value.trim();
  if(!dir){alert('Ingresá una dirección');return;}
  glStatus('Geocodificando...','info');
  // Try multiple geocoding services
  // Method 1: Use Google Maps Geocoding API (works from file://)
  const GMAPS_KEY='AIzaSyDWFzc9CBnY512MK4E48kx7xiDKZI2rv44';
  const q=encodeURIComponent(dir+(ciudad?', '+ciudad:'')+(', Uruguay'));
  const gq=encodeURIComponent(dir+(ciudad?', '+ciudad:'')+', Uruguay');
  const apis=[
    // Google Maps Geocoding (works from file:// via JSONP-style, CORS ok)
    {url:`https://maps.googleapis.com/maps/api/geocode/json?address=${gq}&key=${GMAPS_KEY}&language=es`,
     parse:r=>r.results&&r.results.length?{lat:r.results[0].geometry.location.lat,lon:r.results[0].geometry.location.lng,name:r.results[0].formatted_address}:null},
    // Photon fallback
    {url:`https://photon.komoot.io/api/?q=${q}&limit=1&lang=es`,
     parse:r=>r.features&&r.features.length?{lat:r.features[0].geometry.coordinates[1],lon:r.features[0].geometry.coordinates[0],name:r.features[0].properties.name||dir}:null},
  ];
  let found=false;
  for(const api of apis){
    try{
      const resp=await fetch(api.url,{headers:{'User-Agent':'GuardiaNet/1.0'}});
      const json=await resp.json();
      const result=api.parse(json);
      if(result){
        document.getElementById('gl-lat').value=result.lat.toFixed(7);
        document.getElementById('gl-lon').value=result.lon.toFixed(7);
        glStatus(`✔ Coordenadas: ${result.lat.toFixed(6)}, ${result.lon.toFixed(6)} — ${result.name}`,'success');
        glShowOnMap(result.lat,result.lon,dir);
        found=true;
        break;
      }
    }catch(e){console.warn('Geocoding API failed:',e.message);}
  }
  if(!found){
    glStatus('No se pudo geocodificar. Intentá: 1) Abrí el archivo desde un servidor (python -m http.server) en lugar de doble clic. 2) O ingresá la latitud/longitud manualmente buscando en maps.google.com','error');
  }
}

function glShowOnMap(lat,lon,label){
  const el=document.getElementById('gl-map');
  if(!el)return;
  if(!_glMap){
    el.innerHTML='';
    try{
      _glMap=L.map('gl-map',{zoomControl:true,scrollWheelZoom:false});
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap',maxZoom:18}).addTo(_glMap);
    }catch(e){return;}
  }
  if(_glMarker)_glMap.removeLayer(_glMarker);
  _glMarker=L.marker([lat,lon]).addTo(_glMap).bindPopup('<strong>'+label+'</strong>').openPopup();
  _glMap.setView([lat,lon],15);
}

function glGuardar(){
  const nombre=document.getElementById('gl-nombre').value.trim().toUpperCase();
  const dir=document.getElementById('gl-dir').value.trim();
  const ciudad=document.getElementById('gl-ciudad').value.trim();
  const zona=document.getElementById('gl-zona').value.trim();
  const lat=parseFloat(document.getElementById('gl-lat').value)||null;
  const lon=parseFloat(document.getElementById('gl-lon').value)||null;
  if(!nombre||!dir){alert('Completá al menos el nombre y la dirección');return;}
  const existing=_glLugares.findIndex(l=>l.nombre.toUpperCase()===nombre);
  const entry={nombre,direccion:dir,ciudad,zona,lat,lon};
  if(existing>=0){
    _glLugares[existing]=entry;
    glStatus('✔ Lugar actualizado correctamente.','success');
  } else {
    _glLugares.push(entry);
    glStatus('✔ Lugar agregado correctamente. Total: '+_glLugares.length,'success');
  }
  renderGlTable();
  document.getElementById('gl-list-title').textContent='Lugares registrados ('+_glLugares.length+')';
  // Clear form
  ['gl-nombre','gl-dir','gl-zona','gl-lat','gl-lon'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('gl-ciudad').value='Montevideo';
}

function glStatus(msg, type){
  const el=document.getElementById('gl-status');
  el.style.display='block';
  const colors={info:'var(--blue-light)',success:'var(--green-light)',error:'var(--red-light)'};
  const textColors={info:'var(--blue)',success:'var(--green)',error:'var(--red)'};
  el.style.background=colors[type]||colors.info;
  el.style.color=textColors[type]||textColors.info;
  el.textContent=msg;
}

function glExportarExcel(){
  // Export _glLugares as CSV (simple, no lib needed)
  const header='Nombre,Dirección,Ciudad,Zona,Latitud,Longitud\n';
  const rows=_glLugares.map(l=>[l.nombre,l.direccion,l.ciudad,l.zona,l.lat||'',l.lon||''].map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(',')).join('\n');
  const csv=header+rows;
  const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='lugares_actualizados.csv';a.click();
  glStatus('✔ CSV exportado con '+_glLugares.length+' lugares. Podés importarlo a tu sistema.','success');
}


// ── CARGAR PANEL DE CONTROL ──
let _newPanelData = null;

// Load SheetJS dynamically when needed
function loadSheetJS(callback){
  if(window.XLSX){callback();return;}
  const s=document.createElement('script');
  s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  s.onload=callback;
  s.onerror=()=>uploadStatus('Error cargando librería XLSX. Verificá conexión a internet.','error');
  document.head.appendChild(s);
}

function handleDrop(e){
  e.preventDefault();
  document.getElementById('drop-zone').style.borderColor='var(--border)';
  const file=e.dataTransfer.files[0];
  if(file)processFile(file);
}
function handleFileSelect(input){
  if(input.files[0])processFile(input.files[0]);
}

function uploadStatus(msg, type){
  const el=document.getElementById('upload-status');
  el.style.display='block';
  const bg={info:'var(--blue-light)',success:'var(--green-light)',error:'var(--red-light)',warn:'var(--amber-light)'};
  const col={info:'var(--blue)',success:'var(--green)',error:'var(--red)',warn:'var(--amber)'};
  el.style.background=bg[type]||bg.info;
  el.style.color=col[type]||col.info;
  el.style.border=`1px solid ${col[type]||col.info}40`;
  el.textContent=msg;
}
function setProgress(pct, text){
  document.getElementById('upload-progress').style.display='block';
  document.getElementById('progress-bar').style.width=pct+'%';
  document.getElementById('progress-text').textContent=text;
}

function processFile(file){
  if(!file.name.match(/\.xlsx?$/i)){uploadStatus('El archivo debe ser .xlsx o .xls','error');return;}
  uploadStatus('Cargando '+file.name+'...','info');
  setProgress(10,'Leyendo archivo...');
  loadSheetJS(()=>{
    const reader=new FileReader();
    reader.onload=(e)=>{
      try{
        setProgress(30,'Parseando Excel...');
        const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{raw:false,dateNF:'DD/MM/YYYY'});
        setProgress(50,'Procesando datos...');
        
        // Validate columns
        if(!rows.length){uploadStatus('El archivo está vacío','error');return;}
        const cols=Object.keys(rows[0]);
        const required=['Nombre','Local','Dirección','Entrada planificada','Salida planificada','Fecha'];
        const missing=required.filter(c=>!cols.some(k=>k.toLowerCase().includes(c.toLowerCase())));
        if(missing.length){uploadStatus('Columnas faltantes: '+missing.join(', '),'error');return;}
        
        setProgress(70,'Analizando rutinas...');
        const result=processPanelData(rows);
        setProgress(90,'Preparando resumen...');
        _newPanelData=result;
        
        document.getElementById('upload-info').style.display='block';
        document.getElementById('upload-summary').innerHTML=`
          <strong>Archivo:</strong> ${file.name}<br>
          <strong>Total registros:</strong> ${rows.length.toLocaleString()}<br>
          <strong>Personas únicas:</strong> ${result.stats.totalPersonas}<br>
          <strong>Menos de 44h:</strong> ${result.stats.menos44} personas<br>
          <strong>Con 44h o más:</strong> ${result.stats.mas44} personas<br>
          <strong>Semana detectada:</strong> ${result.stats.semana}
        `;
        setProgress(100,'Listo para aplicar');
        uploadStatus('✔ Archivo procesado. Hacé clic en "Aplicar y actualizar el panel" para cargar los datos.','success');
      }catch(err){
        uploadStatus('Error procesando el archivo: '+err.message,'error');
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function processPanelData(rows){
  const DIAS=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
  const VIDRIEROS=['cristian gabriel machado alvarez','daniel scafarelli','hernan alejandro della nave',
                   'luis gustavo rodriguez barros','miguel amestoy','walter benjamin rivero hernandez'];
  
  function parseH(t){try{const p=String(t).trim().split(':');return parseInt(p[0])+parseInt(p[1])/60;}catch{return null;}}
  function getDia(dateStr){
    try{
      // Handle DD/MM/YYYY or YYYY-MM-DD or JS Date string
      let d;
      if(typeof dateStr==='object')d=dateStr;
      else if(/^\d{4}-/.test(dateStr))d=new Date(dateStr);
      else{const[dd,mm,yy]=dateStr.split('/');d=new Date(yy,mm-1,dd);}
      return DIAS[d.getDay()===0?6:d.getDay()-1];
    }catch{return null;}
  }
  function isExcl(row){
    const l=String(row['Local']||'').toLowerCase();
    const c=String(row['Categoría']||'').toLowerCase();
    const n=String(row['Nombre']||'').toLowerCase();
    const d=String(row['Dirección']||'').toLowerCase();
    if(l.includes('onesta adm'))return'Admin';
    if(c.includes('vidriero')||c.includes('limpia vidrio')||VIDRIEROS.some(v=>n.includes(v)))return'Vidriero';
    if(l.includes('kopel maldonado'))return'Maldonado';
    if(d.includes('maldonado 228')&&!d.includes('montevideo'))return'Maldonado';
    return null;
  }
  
  // Parse all rows
  const operativos=[];
  const fechas=[];
  rows.forEach(row=>{
    if(String(row['Nombre']||'').match(/A - CUBRIR|A-CUBRIR/i))return;
    if(isExcl(row))return;
    const dia=getDia(row['Fecha']);
    if(!dia)return;
    const eh=parseH(row['Entrada planificada']);
    const sh=parseH(row['Salida planificada']);
    let dur=(sh||0)-(eh||0);
    if(dur<0)dur+=24;
    if(dur<=0)return;
    operativos.push({...row, _dia:dia, _eh:eh, _sh:sh, _dur:dur});
    if(row['Fecha'])fechas.push(String(row['Fecha']));
  });
  
  // Hours per person
  const horasMap={};
  operativos.forEach(r=>{
    const n=r['Nombre'];
    horasMap[n]=(horasMap[n]||0)+r._dur;
  });
  
  // Detect week
  const semana=fechas.length?`${[...new Set(fechas)].sort()[0]} — ${[...new Set(fechas)].sort().slice(-1)[0]}`:'Desconocida';
  
  const menos44=Object.entries(horasMap).filter(([,h])=>h<44).map(([n,h])=>({Nombre:n,total_horas:Math.round(h*10)/10})).sort((a,b)=>a.total_horas-b.total_horas);
  const mas44=Object.entries(horasMap).filter(([,h])=>h>=44).map(([n,h])=>({Nombre:n,total_horas:Math.round(h*10)/10})).sort((a,b)=>b.total_horas-a.total_horas);
  
  // Build schedules for menos44
  const schedules={};
  const menos44set=new Set(menos44.map(p=>p.Nombre));
  operativos.filter(r=>menos44set.has(r['Nombre'])).forEach(r=>{
    const n=r['Nombre'],dia=r._dia;
    if(!schedules[n])schedules[n]={};
    if(!schedules[n][dia])schedules[n][dia]=[];
    schedules[n][dia].push({l:r['Local'],e:r['Entrada planificada'],s:r['Salida planificada'],eh:r._eh,sh:r._sh,h:r._dur});
  });
  
  // Build enriched menos44
  const menos44_enriched=menos44.map(p=>{
    const sched=schedules[p.Nombre]||{};
    const dias=Object.keys(sched).sort((a,b)=>DIAS.indexOf(a)-DIAS.indexOf(b));
    const svcs_grp={};
    Object.entries(sched).forEach(([dia,slots])=>{
      slots.forEach(s=>{
        const k=`${s.l}|${s.e}|${s.s}`;
        if(!svcs_grp[k])svcs_grp[k]={local:s.l,entrada:s.e,salida:s.s,horas_dia:s.h,dias:[],zona:''};
        if(!svcs_grp[k].dias.includes(dia))svcs_grp[k].dias.push(dia);
      });
    });
    const servicios=Object.values(svcs_grp).map(s=>({...s,horas_semana:Math.round(s.horas_dia*s.dias.length*10)/10}));
    return {nombre:p.Nombre,horas:p.total_horas,horas_falt:Math.round((44-p.total_horas)*10)/10,dias,zonas:[],servicios};
  });
  
  const stats={
    totalPersonas:Object.keys(horasMap).length,
    menos44:menos44.length, mas44:mas44.length, semana
  };
  
  return {menos44,mas44,schedules,menos44_enriched,stats,
          menos44set:[...menos44set],horasMap};
}

function applyNewPanel(){
  if(!_newPanelData){return;}
  uploadStatus('Aplicando... actualizá la página si los datos no cambian.','info');
  try{
    // Update the core data in D
    D.menos44=_newPanelData.menos44;
    D.mas44=_newPanelData.mas44;
    D.schedules=_newPanelData.schedules;
    D.menos44_enriched=_newPanelData.menos44_enriched;
    D.resumen.menos44=_newPanelData.stats.menos44;
    D.resumen.mas44=_newPanelData.stats.mas44;
    D.resumen.total_personas=_newPanelData.stats.totalPersonas;
    
    // Re-build SCHED
    Object.entries(D.schedules||{}).forEach(([nombre,sched])=>{
      SCHED[nombre]={};
      Object.entries(sched).forEach(([dia,slots])=>{
        SCHED[nombre][dia]=(slots||[]).map(s=>({local:s.l||s.local,entrada:s.e||s.entrada,salida:s.s||s.salida,horas:s.h||s.horas,entrada_h:s.eh||s.entrada_h,salida_h:s.sh||s.salida_h}));
      });
    });
    
    // Rebuild personas_simple
    const DIAS=['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
    D.personas_simple=_newPanelData.menos44_enriched.map(p=>({
      nombre:p.nombre,horas_act:p.horas,horas_falt:p.horas_falt,dias:p.dias,zonas:p.zonas||[]
    })).concat(_newPanelData.mas44.map(p=>({
      nombre:p.Nombre,horas_act:p.total_horas,horas_falt:0,dias:[],zonas:[]
    })));
    
    // Re-render all tabs
    rDash();rM44();rMas44();rPers();populateSel();
    
    saveToSupabase(_newPanelData);
    uploadStatus('✔ Panel actualizado correctamente. Navegá entre las pestañas para ver los datos nuevos.','success');
    showTab('dashboard', document.querySelector('.nav-item'));
    document.querySelectorAll('.nav-item')[0].classList.add('active');
  }catch(e){
    uploadStatus('Error al aplicar: '+e.message,'error');
    console.error(e);
  }
}

// ── INIT ──
rDash();rM44();rMas44();rExcl();rPers();populateSel();rComercial();initDispSelects();initBuscarServicio();initGestorLugares();
document.getElementById('app').style.display='flex';
document.getElementById('pwd-screen').style.display='none';
