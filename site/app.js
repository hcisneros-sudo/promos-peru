const $ = (id) => document.getElementById(id);
const FAVORITES_KEY = "promos-peru-favorites-v1";
const state = { promotions: [], filtered: [], visible: 60, user: null, map: null, layer: null, deferredInstall: null, view: "list", favorites: new Set() };
const bankClass = {"Banco Falabella":"falabella","Banco Ripley":"ripley","Interbank":"interbank","Tarjeta Cencosud":"cencosud","BanBif":"banbif","SIP":"sip"};
const colors = {"Banco Falabella":"#16803f","Banco Ripley":"#6f3dc4","Interbank":"#087f5b","Tarjeta Cencosud":"#df1f37","BanBif":"#e88613","SIP":"#1769aa"};
const controls = ["query","bank","category","subcategory","district","status"];
const escapeHtml = (v="") => String(v).replace(/[&<>'"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
const clean = (v) => v == null ? "" : String(v).trim();
const compact = (v, n=170) => clean(v).length > n ? `${clean(v).slice(0,n).trim()}…` : clean(v);
const formatDate = (v) => /^20\d\d-\d\d-\d\d/.test(clean(v)) ? new Intl.DateTimeFormat("es-PE",{day:"numeric",month:"short",year:"numeric"}).format(new Date(`${v.slice(0,10)}T12:00:00`)) : clean(v);
const distance = (a,b) => { const r=x=>x*Math.PI/180,dLat=r(b[0]-a[0]),dLng=r(b[1]-a[1]); return 12742*Math.asin(Math.sqrt(Math.sin(dLat/2)**2+Math.cos(r(a[0]))*Math.cos(r(b[0]))*Math.sin(dLng/2)**2)); };
function loadFavorites(){ try { const saved=JSON.parse(localStorage.getItem(FAVORITES_KEY)||"[]"); state.favorites=new Set(Array.isArray(saved)?saved.map(String):[]); } catch { state.favorites=new Set(); } }
function saveFavorites(){ try { localStorage.setItem(FAVORITES_KEY,JSON.stringify([...state.favorites])); } catch {} }
function updateFavoriteCount(){ $("favoriteCount").textContent=state.favorites.size.toLocaleString("es-PE"); }
function toggleFavorite(id){ id=String(id);if(state.favorites.has(id))state.favorites.delete(id);else state.favorites.add(id);saveFavorites();updateFavoriteCount();applyFilters(); }

async function loadCatalog(){
  const manifest = await fetch("./data/catalog-manifest.json", {cache:"no-cache"}).then(r => { if(!r.ok) throw new Error("No se encontró el catálogo"); return r.json(); });
  const pieces = await Promise.all(manifest.parts.map(p => fetch(`./data/catalog.parts/${p}`).then(r => { if(!r.ok) throw new Error(`Falta ${p}`); return r.arrayBuffer(); })));
  const size = pieces.reduce((n,p)=>n+p.byteLength,0), joined = new Uint8Array(size); let offset=0;
  pieces.forEach(p=>{ joined.set(new Uint8Array(p),offset); offset+=p.byteLength; });
  if (!("DecompressionStream" in window)) throw new Error("Actualiza Chrome para abrir el catálogo comprimido.");
  const stream = new Blob([joined]).stream().pipeThrough(new DecompressionStream("gzip"));
  const text = await new Response(stream).text();
  return {data:JSON.parse(text),meta:manifest};
}

function unique(values){ return [...new Set(values.map(clean).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es")); }
function setOptions(id, values, label){ const el=$(id), current=el.value; el.innerHTML=`<option value="">${label}</option>`+values.map(v=>`<option>${escapeHtml(v)}</option>`).join(""); el.value=values.includes(current)?current:""; }
function populateFilters(){
  setOptions("bank",unique(state.promotions.map(p=>p.bank)),"Todos los bancos");
  setOptions("category",unique(state.promotions.map(p=>p.category)),"Todas las categorías");
  setOptions("subcategory",unique(state.promotions.filter(p=>!$("category").value||p.category===$("category").value).map(p=>p.subcategory)),"Todas las subcategorías");
  setOptions("district",unique(state.promotions.flatMap(p=>p.locations||[]).map(l=>l.district).filter(v=>v!=="Sin distrito")),"Todos los distritos");
  setOptions("status",unique(state.promotions.map(p=>p.status)),"Todos los estados");
}
function matchingLocations(p){ const district=$("district").value; return (p.locations||[]).filter(l=>!district||l.district===district); }
function applyFilters(){
  const q=$("query").value.toLocaleLowerCase("es").trim(), bank=$("bank").value, category=$("category").value, subcategory=$("subcategory").value, district=$("district").value, status=$("status").value;
  state.filtered=state.promotions.filter(p=>{
    const haystack=[p.name,p.merchant,p.detail,p.benefit,p.category,p.subcategory,...(p.locations||[]).flatMap(l=>[l.address,l.district,l.province,l.department])].join(" ").toLocaleLowerCase("es");
    return (!bank||p.bank===bank)&&(!category||p.category===category)&&(!subcategory||p.subcategory===subcategory)&&(!status||p.status===status)&&(!district||(p.locations||[]).some(l=>l.district===district))&&(!q||haystack.includes(q))&&(state.view!=="favorites"||state.favorites.has(String(p.id)));
  });
  if(state.user) state.filtered.sort((a,b)=>nearestDistance(a)-nearestDistance(b));
  state.visible=60; render();
}
function nearestDistance(p){ const points=(p.locations||[]).filter(l=>Number.isFinite(l.lat)&&Number.isFinite(l.lng)); return points.length?Math.min(...points.map(l=>distance(state.user,[l.lat,l.lng]))):Infinity; }
function card(p){
  const locations=matchingLocations(p), mapped=locations.filter(l=>Number.isFinite(l.lat)&&Number.isFinite(l.lng)), near=state.user&&mapped.length?Math.min(...mapped.map(l=>distance(state.user,[l.lat,l.lng]))):null;
  const place=locations.length>1?`${locations.length} locales`:locations[0]?.district||p.department||"Online o nacional";
  const favorite=state.favorites.has(String(p.id));
  return `<article class="card" data-id="${escapeHtml(p.id)}"><div class="card-strip bank-${bankClass[p.bank]||"default"}"><span>${escapeHtml(p.bank)}</span><span class="card-strip-end"><span>${escapeHtml(p.category)}</span><button class="favorite-button${favorite?" active":""}" data-favorite-id="${escapeHtml(p.id)}" aria-label="${favorite?"Quitar de":"Agregar a"} favoritos" aria-pressed="${favorite}">${favorite?"♥":"♡"}</button></span></div>${p.image?.startsWith("http")?`<img class="card-image" src="${escapeHtml(p.image)}" alt="" loading="lazy" onerror="this.remove()">`:""}<div class="card-body"><p class="merchant">${escapeHtml(p.merchant||p.name)}</p><h2>${escapeHtml(p.benefit||p.discount||p.name)}</h2><p class="description">${escapeHtml(compact(p.detail||p.name))}</p><div class="meta"><span>⌖ ${escapeHtml(place)}${near!=null?` · ${near.toFixed(1)} km`:""}</span>${p.endDate?`<span>◷ Hasta ${escapeHtml(formatDate(p.endDate))}</span>`:""}</div><div class="detail-link">Ver detalles →</div></div></article>`;
}
function render(){
  const locations=state.filtered.flatMap(matchingLocations), mapped=locations.filter(l=>Number.isFinite(l.lat)&&Number.isFinite(l.lng));
  $("resultCount").textContent=`${state.filtered.length.toLocaleString("es-PE")} promociones`;
  $("mapCount").textContent=`${mapped.length.toLocaleString("es-PE")} locales en mapa`;
  const mapMode=state.view==="map";
  $("loading").hidden=true; $("listView").hidden=mapMode; $("mapView").hidden=!mapMode;
  $("listView").innerHTML=state.filtered.slice(0,state.visible).map(card).join("")||`<div class="state">${state.view==="favorites"?"Aún no guardaste promociones favoritas.":"No encontramos promociones con esos filtros."}</div>`;
  $("moreButton").hidden=mapMode||state.visible>=state.filtered.length;
  document.querySelectorAll(".card").forEach(el=>el.addEventListener("click",()=>openDetail(state.promotions.find(p=>p.id===el.dataset.id))));
  document.querySelectorAll(".favorite-button").forEach(el=>el.addEventListener("click",e=>{e.stopPropagation();toggleFavorite(el.dataset.favoriteId);}));
  if(mapMode)renderMap();
}
function openDetail(p){ if(!p)return; const locs=matchingLocations(p); $("detailContent").innerHTML=`<header class="detail-header bank-${bankClass[p.bank]||"default"}"><small>${escapeHtml(p.bank)} · ${escapeHtml(p.category)}${p.subcategory?` · ${escapeHtml(p.subcategory)}`:""}</small><h2>${escapeHtml(p.merchant||p.name)}</h2></header><div class="detail-content"><div class="benefit">${escapeHtml(p.benefit||p.discount||p.name)}</div><h3>Descripción</h3><p>${escapeHtml(p.detail||"El banco no publicó una descripción adicional.")}</p><h3>Dónde aplica</h3><p>${escapeHtml(locs.slice(0,12).map(l=>l.address||l.district).filter(Boolean).join("\n")||"Promoción online o de alcance nacional.")}${locs.length>12?`\n… y ${locs.length-12} locales adicionales.`:""}</p>${p.paymentMethod?`<h3>Medio de pago</h3><p>${escapeHtml(p.paymentMethod)}</p>`:""}${p.endDate?`<h3>Vigencia</h3><p>Hasta ${escapeHtml(formatDate(p.endDate))}</p>`:""}${p.restrictions?`<h3>Restricciones</h3><p>${escapeHtml(p.restrictions)}</p>`:""}${p.terms?`<h3>Términos y condiciones</h3><p>${escapeHtml(p.terms)}</p>`:""}${p.link?`<a class="official" href="${escapeHtml(p.link)}" target="_blank" rel="noopener">Abrir promoción oficial ↗</a>`:""}</div>`; $("detailDialog").showModal(); }
function ensureMap(){ if(state.map)return true; if(typeof L==="undefined"){ $("mapError").hidden=false; $("mapError").textContent="No se pudo iniciar el mapa. Actualiza la aplicación y vuelve a intentarlo."; return false; } $("mapError").hidden=true; state.map=L.map("map",{preferCanvas:true}).setView([-12.08,-77.03],11); L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{attribution:"© OpenStreetMap",maxZoom:19}).addTo(state.map); state.layer=L.layerGroup().addTo(state.map); return true; }
function renderMap(){ if(!ensureMap())return; state.layer.clearLayers(); const points=[]; for(const p of state.filtered){ for(const l of matchingLocations(p)){ if(Number.isFinite(l.lat)&&Number.isFinite(l.lng))points.push({p,l}); } } const shown=points.slice(0,1200); const renderer=L.canvas(); shown.forEach(({p,l})=>L.circleMarker([l.lat,l.lng],{radius:7,color:"#fff",weight:2,fillColor:colors[p.bank]||"#315568",fillOpacity:.93,renderer}).bindPopup(`<b>${escapeHtml(p.merchant||p.name)}</b><br>${escapeHtml(compact(p.benefit||p.detail,100))}<br><small>${escapeHtml(l.address||l.district)}</small>`).addTo(state.layer)); if(state.user)L.circleMarker(state.user,{radius:9,color:"#08263b",fillColor:"#12b99a",fillOpacity:1,renderer}).bindPopup("Tu ubicación").addTo(state.layer); $("mapLimit").textContent=points.length>shown.length?`Mostrando 1,200 de ${points.length.toLocaleString("es-PE")} locales. Usa los filtros para precisar el mapa.`:""; setTimeout(()=>{state.map.invalidateSize();if(shown.length){const bounds=L.latLngBounds(shown.map(({l})=>[l.lat,l.lng]));state.map.fitBounds(bounds,{padding:[24,24],maxZoom:14});}},100); }
function switchView(view){ state.view=view; $("listTab").classList.toggle("active",view==="list"); $("mapTab").classList.toggle("active",view==="map"); $("favoritesTab").classList.toggle("active",view==="favorites"); applyFilters(); }
function bind(){ controls.forEach(id=>$(id).addEventListener(id==="query"?"input":"change",()=>{ if(id==="category")populateFilters(); applyFilters(); })); $("clearButton").onclick=()=>{controls.forEach(id=>$(id).value="");state.user=null;$("message").textContent="";populateFilters();applyFilters();}; $("moreButton").onclick=()=>{state.visible+=60;render();}; $("listTab").onclick=()=>switchView("list"); $("mapTab").onclick=()=>switchView("map"); $("favoritesTab").onclick=()=>switchView("favorites"); $("closeDialog").onclick=()=>$("detailDialog").close(); $("detailDialog").addEventListener("click",e=>{if(e.target===$("detailDialog"))$("detailDialog").close();}); $("nearButton").onclick=()=>navigator.geolocation?navigator.geolocation.getCurrentPosition(({coords})=>{state.user=[coords.latitude,coords.longitude];$("message").textContent="Promociones ordenadas desde la ubicación más cercana.";applyFilters();},()=>$("message").textContent="No se pudo obtener tu ubicación. Revisa el permiso del navegador.",{enableHighAccuracy:true,timeout:12000}):$("message").textContent="Tu navegador no permite usar la ubicación."; }

window.addEventListener("beforeinstallprompt",e=>{e.preventDefault();state.deferredInstall=e;$("installButton").hidden=false;});
$("installButton").onclick=async()=>{if(!state.deferredInstall)return;state.deferredInstall.prompt();await state.deferredInstall.userChoice;state.deferredInstall=null;$("installButton").hidden=true;};
if("serviceWorker" in navigator) window.addEventListener("load",()=>navigator.serviceWorker.register("./sw.js"));

loadFavorites();updateFavoriteCount();bind();
loadCatalog().then(({data,meta})=>{state.promotions=data;populateFilters();applyFilters();$("catalogStatus").textContent=`${data.length.toLocaleString("es-PE")} promociones · Actualizado ${meta.generatedAt.slice(0,10)}`;}).catch(err=>{$("loading").textContent=`No se pudo abrir el catálogo: ${err.message}`;$("catalogStatus").textContent="Catálogo no disponible";});
