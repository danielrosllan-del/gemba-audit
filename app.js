
const LS = {
  get(k,def){ try{ const v = localStorage.getItem(k); return v? JSON.parse(v): def; }catch(e){ return def } },
  set(k,v){ localStorage.setItem(k, JSON.stringify(v)); },
  uid(){ return crypto.randomUUID ? crypto.randomUUID() : String(Date.now())+'-'+Math.random().toString(16).slice(2); }
};

const Enums = {
  estados: ["EJECUTADO","RETRASADO","EN_PROCESO","NO_INICIADO","NO_PROGRAMADO","ELIMINADO"]
};

const App = {
  state: {
    user: null,
    auditorias: LS.get("auditorias", []),
    items: LS.get("aud_items", []),
    planes: LS.get("planes", []),
    estandares: LS.get("estandares", DATA.estandares),
    usuarios: LS.get("usuarios", []),
  },
  save(){
    LS.set("auditorias", this.state.auditorias);
    LS.set("aud_items", this.state.items);
    LS.set("planes", this.state.planes);
    LS.set("estandares", this.state.estandares);
    LS.set("usuarios", this.state.usuarios);
  }
};

const UI = {
  qs: (s)=>document.querySelector(s),
  qsa: (s)=>Array.from(document.querySelectorAll(s)),
  fillSelect(sel, items){ sel.innerHTML=""; items.forEach(v=>{ const o=document.createElement('option'); o.value=v.value??v; o.textContent=v.text??v; sel.appendChild(o); }); },
  toggle(id, show=true){ const el = document.getElementById(id); if(!el) return; el.classList.toggle("hidden", !show); },
  seccion(id){ ["auditoria","planes","admin"].forEach(x=>UI.toggle(x,false)); UI.toggle(id,true); },
  toast(msg){ alert(msg); }
};

const Login = {
  registrar(){
    const nombre = UI.qs("#nombre").value.trim();
    const correo = UI.qs("#correo").value.trim();
    const rol = UI.qs("#rol").value;
    const area = UI.qs("#area").value;
    if(!nombre || !correo){ UI.toast("Completa nombre y correo"); return; }
    const user = {id: LS.uid(), nombre, correo, rol, area};
    App.state.user = user;
    App.state.usuarios.push(user);
    App.save();
    Login.post();
  },
  post(){
    UI.qs("#userBadge").innerHTML = `<span class="tag">${App.state.user.rol}</span> Hola, ${App.state.user.nombre}`;
    UI.toggle("auth", false);
    UI.toggle("dash", true);
    UI.seccion("auditoria", false);
    Dashboard.render();
    // role buttons
    UI.qs("#btnAdmin").style.display = (App.state.user.rol==="Excelencia Operacional") ? "" : "none";
  },
  logout(){ App.state.user=null; UI.toggle("dash", false); UI.toggle("auth", true); }
};

const Dashboard = {
  init(){
    // seed selects
    UI.fillSelect(UI.qs("#rol"), DATA.roles);
    UI.fillSelect(UI.qs("#area"), DATA.areas);
    UI.fillSelect(UI.qs("#audArea"), DATA.areas);
    // procesos placeholder
    this.renderProcesos();
    // Filtros planes
    UI.fillSelect(UI.qs("#fltEstado"), ["(Todos)", ...Enums.estados]);
    UI.fillSelect(UI.qs("#fltResp"), ["(Todos)"]);
    // Buttons
    document.getElementById("btnNuevaAud").onclick=()=>{ UI.seccion("auditoria"); };
    document.getElementById("audFecha").valueAsDate = new Date();
  },
  render(){
    UI.qs("#kpiAud").textContent = App.state.auditorias.length;
    const obs = App.state.items.filter(x=>x.cumple==="NO").length;
    UI.qs("#kpiObs").textContent = obs;
    UI.qs("#kpiPlanes").textContent = App.state.planes.filter(p=>p.estado!=="EJECUTADO" && p.estado!=="ELIMINADO").length;
    UI.qs("#kpiDone").textContent = App.state.planes.filter(p=>p.estado==="EJECUTADO").length;
    this.renderAudList();
  },
  renderProcesos(){
    const area = UI.qs("#audArea").value || DATA.areas[0];
    const procs = DATA.areaProcesos[area] || ["S","P","Q"];
    const cont = UI.qs("#procSel"); cont.innerHTML="";
    procs.forEach(p=>{
      const tag = document.createElement('span'); tag.className="tag"; tag.textContent = p + " - " + (DATA.procesos.find(z=>z.codigo===p)?.nombre||"");
      cont.appendChild(tag);
    });
  },
  renderAudList(){
    const c = UI.qs("#audCards"); c.innerHTML="";
    App.state.auditorias.slice().reverse().forEach(a=>{
      const div = document.createElement('div'); div.className="card";
      div.innerHTML = `<b>${a.fecha}</b> · Área: ${a.area} · Auditor: ${a.auditor}<br>
      Procesos: ${a.procesos.join(", ")}
      <div style="margin-top:8px"><button class="ghost" onclick="Aud.abrir('${a.id}')">Abrir</button></div>`;
      c.appendChild(div);
    });
  }
};

const Aud = {
  crear(){
    if(!App.state.user || App.state.user.rol!=="Auditor"){ UI.toast("Solo Auditor puede crear auditorías"); return; }
    const fecha = UI.qs("#audFecha").value;
    const area = UI.qs("#audArea").value;
    const procesos = DATA.areaProcesos[area] || [];
    const id = LS.uid();
    const aud = {id, fecha, area, procesos, auditor: App.state.user.nombre};
    App.state.auditorias.push(aud);
    // generar items 3A+2B+1C por proceso
    procesos.forEach(p=>{
      const subset = App.state.estandares.filter(e=>e.area===area && e.proceso===p);
      const pick = (crit, n)=> subset.filter(e=>e.criticidad===crit).sort((a,b)=>a.prioridad-b.prioridad).slice(0,n);
      const picked = [...pick("A",3), ...pick("B",2), ...pick("C",1)];
      picked.forEach(st=>{
        App.state.items.push({
          id: LS.uid(),
          audId: id,
          proceso: p,
          estandarId: st.id,
          codigo: st.codigo,
          nombre: st.nombre,
          criticidad: st.criticidad,
          cumple: "SI",
          observacion: "",
          foto: ""
        });
      });
    });
    App.save();
    Dashboard.render();
    UI.seccion("auditoria", false);
    UI.toast("Auditoría creada");
  },
  abrir(id){
    const items = App.state.items.filter(x=>x.audId===id);
    const cont = UI.qs("#audList"); // reuse panel to render details
    const aud = App.state.auditorias.find(x=>x.id===id);
    const title = document.createElement('div'); title.className="card";
    title.innerHTML = `<h3>Auditoría ${aud.fecha} - ${aud.area}</h3>`;
    cont.prepend(title);
    items.forEach(it=>{
      const div = document.createElement('div'); div.className="card";
      div.innerHTML = `<div><b>${it.codigo}</b> - ${it.nombre} <span class="badge ${it.criticidad}">${it.criticidad}</span></div>
      <div class="row">
        <div>
          <label>Cumple</label>
          <select onchange="Aud.setCumple('${it.id}', this.value)">
            <option ${it.cumple==="SI"?"selected":""}>SI</option>
            <option ${it.cumple==="NO"?"selected":""}>NO</option>
            <option ${it.cumple==="NA"?"selected":""}>NA</option>
          </select>
        </div>
        <div>
          <label>Observación (si NO)</label>
          <input value="${it.observacion||""}" onchange="Aud.setObs('${it.id}', this.value)" placeholder="Describa la no conformidad">
        </div>
        <div>
          <label>Foto (URL)</label>
          <input value="${it.foto||""}" onchange="Aud.setFoto('${it.id}', this.value)" placeholder="https://...">
        </div>
      </div>
      <div style="margin-top:8px"><button onclick="Aud.crearPlan('${it.id}')">Crear/Ver Plan</button></div>`;
      UI.qs("#audCards").appendChild(div);
    });
  },
  setCumple(id, v){ const it = App.state.items.find(x=>x.id===id); it.cumple=v; App.save(); },
  setObs(id, v){ const it = App.state.items.find(x=>x.id===id); it.observacion=v; App.save(); },
  setFoto(id, v){ const it = App.state.items.find(x=>x.id===id); it.foto=v; App.save(); },
  crearPlan(itemId){
    const it = App.state.items.find(x=>x.id===itemId);
    let plan = App.state.planes.find(p=>p.itemId===itemId);
    if(!plan){
      plan = {
        id: LS.uid(),
        itemId,
        titulo: `Plan para ${it.codigo}`,
        descripcion: it.observacion || "",
        respArea: App.state.user?.area || "",
        respAccion: "",
        fecha_inicio: "",
        fecha_cierre: "",
        estado: "NO_PROGRAMADO",
        reprogramado_hasta: "",
        verificado_por: "",
        fecha_verificacion: "",
        comentario_cierre: ""
      };
      App.state.planes.push(plan);
      App.save();
    }
    UI.seccion("planes");
    Planes.render();
  }
};

const Planes = {
  render(){
    const est = UI.qs("#fltEstado").value;
    const resp = UI.qs("#fltResp").value;
    const me = App.state.user;
    // fill responsables dynamic
    const resps = ["(Todos)", ...new Set(App.state.planes.map(p=>p.respAccion).filter(Boolean))];
    UI.fillSelect(UI.qs("#fltResp"), resps);
    const cont = UI.qs("#planCards"); cont.innerHTML="";
    let list = App.state.planes.slice();
    if(est && est!=="(Todos)") list = list.filter(x=>x.estado===est);
    if(resp && resp!=="(Todos)") list = list.filter(x=>x.respAccion===resp);
    list.forEach(p=>{
      const item = App.state.items.find(i=>i.id===p.itemId);
      const div = document.createElement('div'); div.className="card";
      div.innerHTML = `<b>${p.titulo}</b> · ${item?.codigo||""} <span class="badge ${item?.criticidad||'A'}">${item?.criticidad||''}</span>
      <div class="row">
        <div><label>Resp. Área</label><input value="${p.respArea||''}" onchange="Planes.set('${p.id}','respArea',this.value)" ${me.rol!=="Responsable de Área"?"readonly":""}></div>
        <div><label>Resp. Acción</label><input value="${p.respAccion||''}" onchange="Planes.set('${p.id}','respAccion',this.value)" ${me.rol!=="Responsable de Área"?"readonly":""}></div>
        <div><label>Inicio</label><input type="date" value="${p.fecha_inicio||''}" onchange="Planes.set('${p.id}','fecha_inicio',this.value)" ${me.rol!=="Responsable de Área"?"readonly":""}></div>
        <div><label>Cierre</label><input type="date" value="${p.fecha_cierre||''}" onchange="Planes.set('${p.id}','fecha_cierre',this.value)" ${me.rol!=="Responsable de Área"?"readonly":""}></div>
        <div><label>Estado</label>
          <select onchange="Planes.set('${p.id}','estado',this.value)">
            ${Enums.estados.map(s=>`<option ${p.estado===s?'selected':''}>${s}</option>`).join('')}
          </select>
        </div>
      </div>
      <div class="row">
        <div><label>Descripción</label><textarea rows="2" onchange="Planes.set('${p.id}','descripcion',this.value)">${p.descripcion||''}</textarea></div>
        <div><label>Evidencia (URL)</label><input onchange="Planes.set('${p.id}','evidencia',this.value)" value="${p.evidencia||''}" placeholder="https://foto.pdf"></div>
      </div>
      <div style="margin-top:8px">
        ${(me.rol==="Auditor")?`<button class="ghost" onclick="Planes.reprogramar('${p.id}')">Reprogramar (Auditor)</button>`:""}
        ${(me.rol==="Auditor")?`<button onclick="Planes.verificar('${p.id}')">Verificar y Cerrar</button>`:""}
      </div>`;
      cont.appendChild(div);
    });
  },
  set(id, field, val){
    const p = App.state.planes.find(x=>x.id===id);
    // Permisos básicos
    const me = App.state.user;
    if(["respArea","respAccion","fecha_inicio","fecha_cierre"].includes(field) && me.rol!=="Responsable de Área"){
      UI.toast("Solo Responsable de Área puede asignar responsables y fechas");
      return;
    }
    p[field]=val; App.save();
  },
  reprogramar(id){
    const me = App.state.user;
    if(me.rol!=="Auditor"){ UI.toast("Solo Auditor reprograma"); return; }
    const p = App.state.planes.find(x=>x.id===id);
    const hasta = prompt("Nueva fecha de cierre (YYYY-MM-DD):", p.fecha_cierre||"");
    if(hasta){
      p.reprogramado_hasta = hasta;
      p.fecha_cierre = hasta;
      App.save();
      this.render();
    }
  },
  verificar(id){
    const me = App.state.user;
    if(me.rol!=="Auditor"){ UI.toast("Solo Auditor verifica"); return; }
    const p = App.state.planes.find(x=>x.id===id);
    p.estado = "EJECUTADO";
    p.verificado_por = me.nombre;
    p.fecha_verificacion = new Date().toISOString().slice(0,10);
    App.save();
    this.render();
  }
};

const AdminStd = {
  guardar(){
    const area = UI.qs("#admArea").value;
    const proc = UI.qs("#admProc").value;
    const crit = UI.qs("#admCrit").value;
    const codigo = UI.qs("#admCod").value;
    const nombre = UI.qs("#admNom").value;
    const url = UI.qs("#admUrl").value;
    const pri = parseInt(UI.qs("#admPri").value||"1",10);
    App.state.estandares.push({id:LS.uid(), area, proceso:proc, criticidad:crit, prioridad:pri, codigo, nombre, archivo_url:url});
    App.save();
    this.listar();
  },
  listar(){
    const area = UI.qs("#admArea").value;
    const proc = UI.qs("#admProc").value;
    const cont = UI.qs("#stdList"); cont.innerHTML="";
    const list = App.state.estandares.filter(e=>e.area===area && e.proceso===proc).sort((a,b)=>a.criticidad.localeCompare(b.criticidad) || a.prioridad-b.prioridad);
    list.forEach(e=>{
      const d = document.createElement('div'); d.className="card";
      d.innerHTML = `<b>${e.codigo}</b> - ${e.nombre} <span class="badge ${e.criticidad}">${e.criticidad}</span> · prioridad ${e.prioridad} ${e.archivo_url?`· <a href="${e.archivo_url}" target="_blank">archivo</a>`:""}`;
      cont.appendChild(d);
    });
  }
};

window.addEventListener("load", ()=>{
  // PWA registration
  if("serviceWorker" in navigator){
    navigator.serviceWorker.register("./service-worker.js");
  }
  Dashboard.init();
  // seed selects for admin
  UI.fillSelect(UI.qs("#admArea"), DATA.areas);
  UI.fillSelect(UI.qs("#admProc"), DATA.procesos.map(p=>({value:p.codigo,text:`${p.codigo} - ${p.nombre}`})));
  // change handlers
  UI.qs("#audArea").addEventListener("change", ()=>Dashboard.renderProcesos());
  // if returning user
  const users=App.state.usuarios;
  if(App.state.user){ Login.post(); } else if(users.length){
    UI.qs("#nombre").value=users[0].nombre; UI.qs("#correo").value=users[0].correo;
  }
  Dashboard.render();
});
