// ---- Instellingen ----
// Bestand met de gemeentelijst. Kolom A = gemeente, kolom B = "Ja" bij actieve gemeentes.
const LIST_FILE  = "actieve-gemeentes.xlsx";
// Naam van het werkblad in dat bestand (bestaat het niet, dan wordt het eerste werkblad gebruikt)
const LIST_SHEET = "Actieve gemeentes";

let ACTIVE = [];

// Alternatieve schrijfwijzen (weergavenaam -> officiële naam)
const ALIAS = {"Den Haag":"'s-Gravenhage"};

// Bron van de grenzen: eerst lokaal bestand (naast dit HTML-bestand), anders CBS-data via cartomap
const SOURCES = ["gemeente_2025.topojson","https://cartomap.github.io/nl/wgs84/gemeente_2025.topojson"];

const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g,"");
const stripParen = s => s.replace(/\s*\(.*?\)\s*/g,"");


// ---- Lijst met gemeentes inlezen uit Excel ----
// Alleen rijen met "Ja" in kolom B worden gemarkeerd (heeft het werkblad maar één kolom,
// dan worden alle gemeentes gemarkeerd). Kopregels, het voorvoegsel "Gemeente " en
// dubbele regels worden automatisch genegeerd.
const YES = ["ja","j","yes","y","x","1","true","waar"];
let LIST_INFO = "";

function cleanList(rows){          // rows = [[naam, actief], ...]
  const hasFlags = rows.some(r => String(r[1] ?? "").trim() !== "");
  const seen = new Map();
  rows.forEach(r=>{
    let n = String(r[0] ?? "").trim().replace(/^gemeente\s+/i,"");
    if(!n || /^gemeente$/i.test(n)) return;
    if(hasFlags && !YES.includes(String(r[1] ?? "").trim().toLowerCase())) return;
    const k = norm(n);
    if(!seen.has(k)) seen.set(k, n);
  });
  return [...seen.values()];
}

async function loadList(){
  if(typeof XLSX==="undefined") throw new Error("De Excel-bibliotheek kon niet worden geladen.");
  const r = await fetch(LIST_FILE);
  if(!r.ok) throw new Error("Het bestand \""+LIST_FILE+"\" is niet gevonden. Zet het in dezelfde map als deze pagina.");
  const wb = XLSX.read(await r.arrayBuffer());
  const sheet = wb.SheetNames.includes(LIST_SHEET) ? LIST_SHEET : wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1});
  const list = cleanList(rows);
  if(!list.length) throw new Error("Geen actieve gemeentes gevonden in werkblad \""+sheet+"\" van \""+LIST_FILE+"\".");
  LIST_INFO = "Bron: "+LIST_FILE+" (werkblad \""+sheet+"\")";
  return list;
}

async function load(){
  for(const url of SOURCES){
    try{ const r = await fetch(url); if(r.ok) return await r.json(); }catch(e){}
  }
  throw new Error("Kon de gemeentegrenzen niet laden.");
}

Promise.all([load(), loadList()]).then(([topo, names])=>{
  ACTIVE = names;
  if(typeof d3==="undefined"||typeof topojson==="undefined") throw new Error("De kaartbibliotheken (d3/topojson) konden niet worden geladen. Controleer je internetverbinding.");
  const key = Object.keys(topo.objects)[0];
  const fc = topojson.feature(topo, topo.objects[key]);
  const wanted = new Map(); // genormaliseerde naam -> weergavenaam
  ACTIVE.forEach(n=>wanted.set(norm(ALIAS[n]||n), n));

  const matchName = f => {
    const n = f.properties.statnaam;
    return wanted.get(norm(n)) || wanted.get(norm(stripParen(n))) || null;
  };

  const svg = d3.select("#svg");
  const proj = d3.geoMercator().fitSize([600,720], fc);
  const path = d3.geoPath(proj);
  const tip = document.getElementById("tip");
  const found = new Set();

  const paths = svg.selectAll("path").data(fc.features).join("path")
    .attr("class", f=>{ const m=matchName(f); if(m) found.add(m); return "g"+(m?" active":""); })
    .attr("d", path)
    .each(function(f){ this.__name = matchName(f); });

  // actieve gemeentes bovenop tekenen
  svg.selectAll("path.active").raise();

  const activeSel = svg.selectAll("path.active");
  activeSel.on("mousemove", function(ev){
      tip.style.display="block"; tip.textContent=this.__name;
      const b = svg.node().getBoundingClientRect(), p = tip.parentNode.getBoundingClientRect();
      tip.style.left = (ev.clientX-p.left+12)+"px"; tip.style.top=(ev.clientY-p.top+12)+"px";
      setOn(this.__name);
    })
    .on("mouseleave", ()=>{ tip.style.display="none"; setOn(null); });

  const list = document.getElementById("list");
  const items = {};
  [...ACTIVE].sort((a,b)=>a.localeCompare(b,"nl")).forEach(n=>{
    const li = document.createElement("li"); li.textContent=n;
    li.onmouseenter=()=>setOn(n); li.onmouseleave=()=>setOn(null);
    list.appendChild(li); items[n]=li;
  });

  function setOn(name){
    activeSel.classed("on", function(){return this.__name===name});
    Object.entries(items).forEach(([n,li])=>li.classList.toggle("on", n===name));
  }

  const count = document.getElementById("count");
  const setCount = t => count.textContent = t;
  setCount(ACTIVE.length+" actieve gemeentes");
  const src = document.createElement("div"); src.className="source"; src.textContent = LIST_INFO;
  count.after(src);

  document.getElementById("q").addEventListener("input", e=>{
    const q = e.target.value.trim().toLowerCase(); let v=0;
    Object.entries(items).forEach(([n,li])=>{
      const m = n.toLowerCase().includes(q); li.classList.toggle("hide",!m); if(m) v++;
    });
    activeSel.classed("dim", function(){ return q && !this.__name.toLowerCase().includes(q); });
    setCount(q ? v+" van "+ACTIVE.length+" gemeentes" : ACTIVE.length+" actieve gemeentes");
  });

  const missing = ACTIVE.filter(n=>!found.has(n));
  if(missing.length) console.warn("Niet gevonden in kaartdata:", missing);
}).catch(err=>{
  const e = document.getElementById("err");
  e.style.display="block";
  e.textContent = "Fout: "+err.message;
  console.error(err);
});
