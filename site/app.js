const $ = id => document.getElementById(id);
const FAVORITES_KEY = "promos-peru-favorites-v1";
const FILTER_KEYS = ["bank", "category", "subcategory", "status", "department", "province", "district", "day"];
const state = {promotions:[],filtered:[],visible:60,user:null,map:null,layer:null,detailMap:null,deferredInstall:null,view:"list",favorites:new Set(),scrollY:0,selected:Object.fromEntries(FILTER_KEYS.map(k=>[k,new Set()])),discountActive:false};
const bankClass={"Banco Falabella":"falabella","Banco Ripley":"ripley","Interbank":"interbank","Tarjeta Cencosud":"cencosud","BanBif":"banbif","SIP":"sip"};
const colors={"Banco Falabella":"#16803f","Banco Ripley":"#6f3dc4","Interbank":"#087f5b","Tarjeta Cencosud":"#df1f37","BanBif":"#e88613","SIP":"#1769aa"};
const escapeHtml=(v="")=>String(v).replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));
const clean=v=>v==null?"":String(v).trim();
const compact=(v,n=170)=>clean(v).length>n?`${clean(v).slice(0,n).trim()}…`:clean(v);
const formatDate=v=>/^20\d\d-\d\d-\d\d/.test(clean(v))?new Intl.DateTimeFormat("es-PE",{day:"numeric",month:"short",year:"numeric"}).format(new Date(`${v.slice(0,10)}T12:00:00`)):clean(v);
const formatMoney=v=>Number.isFinite(Number(v))?`S/${Number(v).toLocaleString("es-PE",{maximumFractionDigits:2})}`:"";
const distance=(a,b)=>{const r=x=>x*Math.PI/180,dLat=r(b[0]-a[0]),dLng=r(b[1]-a[1]);return 12742*Math.asin(Math.sqrt(Math.sin(dLat/2)**2+Math.cos(r(a[0]))*Math.cos(r(b[0]))*Math.sin(dLng/2)**2));};
const unique=values=>[...new Set(values.map(clean).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));
const physical=l=>!l.type||l.type==="local_fisico";

async function loadCatalog(){
  const manifest=await fetch("./data/catalog-manifest.json",{cache:"no-cache"}).then(r=>{if(!r.ok)throw new Error("No se encontró el catálogo");return r.json();});
  const pieces=await Promise.all(manifest.parts.map(p=>fetch(`./data/catalog.parts/${p}`,{cache:"no-cache"}).then(r=>{if(!r.ok)throw new Error(`Falta ${p}`);return r.arrayBuffer();})));
  const size=pieces.reduce((n,p)=>n+p.byteLength,0),joined=new Uint8Array(size);let offset=0;pieces.forEach(p=>{joined.set(new Uint8Array(p),offset);offset+=p.byteLength;});
  if(!("DecompressionStream" in window))throw new Error("Actualiza Chrome para abrir el catálogo comprimido.");
  const stream=new Blob([joined]).stream().pipeThrough(new DecompressionStream("gzip"));
  return {data:JSON.parse(await new Response(stream).text()),meta:manifest};
}

function loadFavorites(){try{const saved=JSON.parse(localStorage.getItem(FAVORITES_KEY)||"[]");state.favorites=new Set(Array.isArray(saved)?saved.map(String):[]);}catch{state.favorites=new Set();}}
function saveFavorites(){try{localStorage.setItem(FAVORITES_KEY,JSON.stringify([...state.favorites]));}catch{}}
function updateFavoriteCount(){$("favoriteCount").textContent=state.favorites.size.toLocaleString("es-PE");}
function toggleFavorite(id){id=String(id);state.favorites.has(id)?state.favorites.delete(id):state.favorites.add(id);saveFavorites();updateFavoriteCount();applyFilters();}

function filterValues(key){
  const rows=state.promotions;
  if(key==="bank")return unique(rows.map(p=>p.bank));
  if(key==="category")return unique(rows.map(p=>p.category));
  if(key==="subcategory")return unique(rows.filter(p=>!state.selected.category.size||state.selected.category.has(p.category)).map(p=>p.subcategory));
  if(key==="status")return unique(rows.map(p=>p.status));
  if(key==="day")return unique(rows.flatMap(p=>p.days||[]));
  const locations=rows.flatMap(p=>p.locations||[]).filter(l=>(!state.selected.department.size||state.selected.department.has(clean(l.department)))&&(!state.selected.province.size||state.selected.province.has(clean(l.province))));
  if(key==="department")return unique(rows.flatMap(p=>p.locations||[]).map(l=>l.department));
  if(key==="province")return unique(locations.map(l=>l.province));
  return unique(locations.map(l=>l.district).filter(v=>v!=="Sin distrito"));
}
function renderMulti(key){
  const values=filterValues(key),selected=state.selected[key];for(const value of [...selected])if(!values.includes(value))selected.delete(value);
  const target=$(`${key}Options`);target.innerHTML=values.length?values.map(value=>`<label><input type="checkbox" value="${escapeHtml(value)}" ${selected.has(value)?"checked":""}><span>${escapeHtml(value)}</span></label>`).join(""):`<p>Sin opciones disponibles</p>`;
  target.querySelectorAll("input").forEach(input=>input.addEventListener("change",()=>{input.checked?selected.add(input.value):selected.delete(input.value);if(["category","department","province"].includes(key))populateFilters();updateFilterLabels();applyFilters();}));
}
function updateFilterLabels(){let total=state.discountActive?1:0;for(const key of FILTER_KEYS){const n=state.selected[key].size;total+=n;$(`${key}Label`).textContent=n?`${n} seleccionado${n>1?"s":""}`:"Todos";}$("advancedCount").textContent=total?String(total):"";$("advancedCount").hidden=!total;}
function populateFilters(){FILTER_KEYS.forEach(renderMulti);updateFilterLabels();}
function matchesSet(value,set){return !set.size||set.has(clean(value));}
function locationMatches(l){return matchesSet(l.department,state.selected.department)&&matchesSet(l.province,state.selected.province)&&matchesSet(l.district,state.selected.district);}
function matchingLocations(p){return(p.locations||[]).filter(locationMatches);}
function matchDiscount(p){if(!state.discountActive)return true;const value=Number(p.discount);return Number.isFinite(value)&&value>=Number($("discountMin").value)&&value<=Number($("discountMax").value);}
function applyFilters(){
  const q=$("query").value.toLocaleLowerCase("es").trim(),geoActive=state.selected.department.size||state.selected.province.size||state.selected.district.size;
  state.filtered=state.promotions.filter(p=>{const locations=p.locations||[],haystack=[p.name,p.merchant,p.detail,p.benefit,p.category,p.subcategory,p.paymentMethod,p.validDaysText,...locations.flatMap(l=>[l.address,l.district,l.province,l.department])].join(" ").toLocaleLowerCase("es");return matchesSet(p.bank,state.selected.bank)&&matchesSet(p.category,state.selected.category)&&matchesSet(p.subcategory,state.selected.subcategory)&&matchesSet(p.status,state.selected.status)&&(!state.selected.day.size||(p.days||[]).some(day=>state.selected.day.has(day)))&&(!geoActive||locations.some(locationMatches))&&matchDiscount(p)&&(!q||haystack.includes(q))&&(state.view!=="favorites"||state.favorites.has(String(p.id)));});
  if(state.user)state.filtered.sort((a,b)=>nearestDistance(a)-nearestDistance(b));state.visible=60;render();
}
function nearestDistance(p){const points=(p.locations||[]).filter(l=>physical(l)&&Number.isFinite(l.lat)&&Number.isFinite(l.lng));return points.length?Math.min(...points.map(l=>distance(state.user,[l.lat,l.lng]))):Infinity;}
function card(p){
  const locations=matchingLocations(p).filter(physical),mapped=locations.filter(l=>Number.isFinite(l.lat)&&Number.isFinite(l.lng)),near=state.user&&mapped.length?Math.min(...mapped.map(l=>distance(state.user,[l.lat,l.lng]))):null,place=locations.length>1?`${locations.length} locales`:locations[0]?.district||locations[0]?.department||"Online o nacional",favorite=state.favorites.has(String(p.id));
  return `<article class="card" data-id="${escapeHtml(p.id)}"><div class="card-strip bank-${bankClass[p.bank]||"default"}"><span>${escapeHtml(p.bank)}</span><span class="card-strip-end"><span>${escapeHtml(p.category)}</span><button class="favorite-button${favorite?" active":""}" data-favorite-id="${escapeHtml(p.id)}" aria-label="${favorite?"Quitar de":"Agregar a"} favoritos">${favorite?"♥":"♡"}</button></span></div>${p.image?.startsWith("http")?`<img class="card-image" src="${escapeHtml(p.image)}" alt="" loading="lazy" onerror="this.remove()">`:""}<div class="card-body"><p class="merchant">${escapeHtml(p.merchant||p.name)}</p><h2>${escapeHtml(p.benefit||p.name)}</h2><p class="description">${escapeHtml(compact(p.detail||p.name))}</p><div class="meta"><span>⌖ ${escapeHtml(place)}${near!=null?` · ${near.toFixed(1)} km`:""}</span>${p.endDate?`<span>◷ Hasta ${escapeHtml(formatDate(p.endDate))}</span>`:""}${p.days?.length?`<span>◫ ${escapeHtml(p.days.join(", "))}</span>`:""}</div><div class="detail-link">Ver detalles →</div></div></article>`;
}
function render(){
  const locations=state.filtered.flatMap(matchingLocations).filter(physical),mapped=locations.filter(l=>Number.isFinite(l.lat)&&Number.isFinite(l.lng));$("resultCount").textContent=`${state.filtered.length.toLocaleString("es-PE")} promociones`;$("mapCount").textContent=`${mapped.length.toLocaleString("es-PE")} locales en mapa`;
  const mapMode=state.view==="map";$("loading").hidden=true;$("listView").hidden=mapMode;$("mapView").hidden=!mapMode;$("listView").innerHTML=state.filtered.slice(0,state.visible).map(card).join("")||`<div class="state">${state.view==="favorites"?"Aún no guardaste promociones favoritas.":"No encontramos promociones con esos filtros."}</div>`;$("moreButton").hidden=mapMode||state.visible>=state.filtered.length;
  document.querySelectorAll(".card").forEach(el=>el.addEventListener("click",()=>openDetail(state.promotions.find(p=>String(p.id)===el.dataset.id),false)));document.querySelectorAll(".favorite-button").forEach(el=>el.addEventListener("click",e=>{e.stopPropagation();toggleFavorite(el.dataset.favoriteId);}));if(mapMode)renderMap();
}

function detailFields(p){const fields=[];if(p.benefitType)fields.push(["Tipo de beneficio",p.benefitType.replaceAll("_"," ")]);if(p.discount!=null)fields.push(["Descuento",`${p.discount}%`]);if(p.days?.length||p.validDaysText)fields.push(["Días válidos",p.days?.join(", ")||p.validDaysText]);if(p.paymentMethod)fields.push(["Medio de pago",p.paymentMethod]);if(p.channels?.length)fields.push(["Canales",p.channels.join(", ")]);if(p.purchaseMinimum!=null)fields.push(["Compra mínima",formatMoney(p.purchaseMinimum)]);if(p.discountCap!=null)fields.push(["Tope de descuento",formatMoney(p.discountCap)]);if(p.requiresCoupon)fields.push(["Cupón","Requerido"]);if(p.requiresReservation)fields.push(["Reserva","Requerida"]);if(p.newCustomersOnly)fields.push(["Clientes","Solo nuevos clientes"]);if(p.subjectToStock)fields.push(["Stock","Sujeto a disponibilidad"]);return fields;}
function lockBackground(){state.scrollY=window.scrollY;document.body.classList.add("modal-open");document.body.style.top=`-${state.scrollY}px`;}
function unlockBackground(){if(!document.body.classList.contains("modal-open"))return;document.body.classList.remove("modal-open");document.body.style.top="";window.scrollTo(0,state.scrollY);}
function closeDetail(){if($("detailDialog").open)$("detailDialog").close();if(state.detailMap){state.detailMap.remove();state.detailMap=null;}unlockBackground();}
function openDetail(p,fromMap=false){
  if(!p)return;const locs=matchingLocations(p).filter(physical),mapped=locs.filter(l=>Number.isFinite(l.lat)&&Number.isFinite(l.lng)),where=locs.slice(0,15).map(l=>[l.name,l.address,l.district,l.department].filter(Boolean).join(" · ")).filter(Boolean),fields=detailFields(p).map(([a,b])=>`<div><dt>${escapeHtml(a)}</dt><dd>${escapeHtml(b)}</dd></div>`).join(""),mapBlock=!fromMap&&mapped.length?`<h3>Mapa de ubicaciones</h3><div class="detail-map-wrap"><div id="detailMap"></div><button id="detailLocateButton" class="locate-map" type="button">◎ Mi ubicación</button></div>`:"";
  $("detailContent").innerHTML=`<header class="detail-header bank-${bankClass[p.bank]||"default"}"><small>${escapeHtml(p.bank)} · ${escapeHtml(p.category)}${p.subcategory?` · ${escapeHtml(p.subcategory)}`:""}</small><h2>${escapeHtml(p.merchant||p.name)}</h2></header><div class="detail-content"><div class="benefit">${escapeHtml(p.benefit||p.name)}</div><h3>Descripción</h3><p>${escapeHtml(p.detail||"El banco no publicó una descripción adicional.")}</p>${fields?`<dl class="detail-fields">${fields}</dl>`:""}<h3>Dónde aplica</h3><p>${escapeHtml(where.join("\n")||"Promoción online o de alcance nacional.")}${locs.length>15?`\n… y ${locs.length-15} locales adicionales.`:""}</p>${p.restrictions?`<h3>Restricciones</h3><p>${escapeHtml(p.restrictions)}</p>`:""}${p.terms?`<h3>Términos y condiciones completos</h3><p class="terms">${escapeHtml(p.terms)}</p>`:""}${mapBlock}${p.link?`<a class="official" href="${escapeHtml(p.link)}" target="_blank" rel="noopener">Abrir promoción oficial ↗</a>`:""}</div>`;
  lockBackground();$("detailDialog").showModal();if(!fromMap&&mapped.length)setTimeout(()=>renderDetailMap(mapped),50);
}

function getUserLocation(onSuccess){if(state.user){onSuccess(state.user);return;}if(!navigator.geolocation){$("message").textContent="Tu navegador no permite usar la ubicación.";return;}navigator.geolocation.getCurrentPosition(({coords})=>{state.user=[coords.latitude,coords.longitude];$("message").textContent="Ubicación encontrada.";onSuccess(state.user);},()=>$("message").textContent="No se pudo obtener tu ubicación. Revisa el permiso del navegador.",{enableHighAccuracy:true,timeout:12000});}
function ensureMap(){if(state.map)return true;if(typeof L==="undefined"){$("mapError").hidden=false;$("mapError").textContent="No se pudo iniciar el mapa. Actualiza la aplicación y vuelve a intentarlo.";return false;}$("mapError").hidden=true;state.map=L.map("map",{preferCanvas:true}).setView([-12.08,-77.03],11);L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{attribution:"© OpenStreetMap",maxZoom:19}).addTo(state.map);state.layer=L.layerGroup().addTo(state.map);return true;}
function addUserMarker(target){if(!state.user)return;L.circleMarker(state.user,{radius:9,color:"#08263b",fillColor:"#12b99a",fillOpacity:1}).bindPopup("Tu ubicación").addTo(target);}
function renderMap(){
  if(!ensureMap())return;state.layer.clearLayers();const points=[];for(const p of state.filtered)for(const l of matchingLocations(p))if(physical(l)&&Number.isFinite(l.lat)&&Number.isFinite(l.lng))points.push({p,l});const renderer=L.canvas();
  points.forEach(({p,l})=>{const popup=`<div class="map-popup"><b>${escapeHtml(p.merchant||p.name)}</b><p>${escapeHtml(compact(p.benefit||p.detail,100))}</p><small>${escapeHtml(l.address||l.district)}</small><button type="button" data-map-detail="${escapeHtml(p.id)}">Ver más</button></div>`;L.circleMarker([l.lat,l.lng],{radius:7,color:"#fff",weight:2,fillColor:colors[p.bank]||"#315568",fillOpacity:.93,renderer}).bindPopup(popup).addTo(state.layer);});
  addUserMarker(state.layer);$("mapLimit").textContent="";setTimeout(()=>{state.map.invalidateSize();if(points.length)state.map.fitBounds(L.latLngBounds(points.map(({l})=>[l.lat,l.lng])),{padding:[24,24],maxZoom:14});},100);
}
function renderDetailMap(mapped){if(state.detailMap)state.detailMap.remove();state.detailMap=L.map("detailMap",{preferCanvas:true,scrollWheelZoom:false});L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{attribution:"© OpenStreetMap",maxZoom:19}).addTo(state.detailMap);mapped.forEach(l=>L.circleMarker([l.lat,l.lng],{radius:7,color:"#fff",weight:2,fillColor:"#0b967d",fillOpacity:.95}).bindPopup(escapeHtml(l.address||l.name||"Local")).addTo(state.detailMap));state.detailMap.fitBounds(L.latLngBounds(mapped.map(l=>[l.lat,l.lng])),{padding:[20,20],maxZoom:15});$("detailLocateButton").onclick=()=>getUserLocation(pos=>{addUserMarker(state.detailMap);state.detailMap.setView(pos,15);});}
function switchView(view){state.view=view;$("listTab").classList.toggle("active",view==="list");$("mapTab").classList.toggle("active",view==="map");$("favoritesTab").classList.toggle("active",view==="favorites");applyFilters();}
function clearFilters(){FILTER_KEYS.forEach(key=>state.selected[key].clear());state.discountActive=false;$("discountMin").value=0;$("discountMax").value=100;$("query").value="";state.user=null;$("message").textContent="";updateDiscount();populateFilters();applyFilters();}
function updateDiscount(){let min=Number($("discountMin").value),max=Number($("discountMax").value);if(min>max){[$("discountMin").value,$("discountMax").value]=[max,min];[min,max]=[max,min];}$("discountValue").textContent=`${min}% – ${max}%`;updateFilterLabels();}
function bind(){
  $("query").addEventListener("input",applyFilters);$("clearButton").onclick=clearFilters;$("moreButton").onclick=()=>{state.visible+=60;render();};$("listTab").onclick=()=>switchView("list");$("mapTab").onclick=()=>switchView("map");$("favoritesTab").onclick=()=>switchView("favorites");
  $("closeDialog").onclick=closeDetail;$("detailDialog").addEventListener("cancel",e=>{e.preventDefault();closeDetail();});$("detailDialog").addEventListener("click",e=>{if(e.target===$("detailDialog"))closeDetail();});
  $("nearButton").onclick=()=>getUserLocation(()=>{$("message").textContent="Promociones ordenadas desde la ubicación más cercana.";applyFilters();if(state.view==="map"){addUserMarker(state.layer);state.map.setView(state.user,14);}});$("mapLocateButton").onclick=()=>getUserLocation(pos=>{renderMap();state.map.setView(pos,14);});
  for(const id of ["discountMin","discountMax"])$(id).addEventListener("input",()=>{state.discountActive=true;updateDiscount();applyFilters();});
  document.addEventListener("click",e=>{const button=e.target.closest("[data-map-detail]");if(button)openDetail(state.promotions.find(p=>String(p.id)===button.dataset.mapDetail),true);});document.addEventListener("click",e=>{if(!e.target.closest(".multi-filter"))document.querySelectorAll(".multi-filter[open]").forEach(d=>d.removeAttribute("open"));});
}
window.addEventListener("beforeinstallprompt",e=>{e.preventDefault();state.deferredInstall=e;$("installButton").hidden=false;});$("installButton").onclick=async()=>{if(!state.deferredInstall)return;state.deferredInstall.prompt();await state.deferredInstall.userChoice;state.deferredInstall=null;$("installButton").hidden=true;};if("serviceWorker" in navigator)window.addEventListener("load",()=>navigator.serviceWorker.register("./sw.js"));
loadFavorites();updateFavoriteCount();bind();updateDiscount();loadCatalog().then(({data,meta})=>{state.promotions=data;populateFilters();applyFilters();$("catalogStatus").textContent=`${data.length.toLocaleString("es-PE")} promociones · Actualizado ${meta.generatedAt.slice(0,10)}`;}).catch(err=>{$("loading").textContent=`No se pudo abrir el catálogo: ${err.message}`;$("catalogStatus").textContent="Catálogo no disponible";});
