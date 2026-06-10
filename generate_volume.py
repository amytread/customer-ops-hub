#!/usr/bin/env python3
"""
Render the Volume Pulse dashboard (customer-pulse/volume-pulse/index.html) from
the three CSVs written by refresh_volume.py. Tread-branded, Chart.js, with a
Daily / Weekly / Monthly toggle. No external data calls — pure CSV → HTML.
"""
import csv, datetime, json, os

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
VOL_DIR = os.path.join(OUT_DIR, "volume-pulse")
TODAY = datetime.date.today().isoformat()

TREAD_WORDMARK = '''<svg xmlns="http://www.w3.org/2000/svg" width="120" height="24" viewBox="0 0 150 30" fill="none">
<path d="M59.0903 4.36316H73.9359V8.54708H68.92V25.2795H64.1078V8.54547H59.0903V4.36316Z" fill="white"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M90.856 10.5747C90.856 6.46764 87.8866 4.36292 83.9704 4.36292H76.4194V25.2793H81.2317V17.0138L87.6813 25.2777H93.8L86.5308 16.6346C89.1151 15.823 90.856 13.7182 90.856 10.5747ZM85.7863 10.9077C85.7863 13.5198 83.611 13.6463 81.7937 13.6463H81.2304V8.16905H81.7937C83.611 8.16905 85.7863 8.29559 85.7863 10.9077Z" fill="white"/>
<path d="M107.853 8.54522H100.582V12.7067H107.544V16.8906H100.582V21.0954H107.851V25.2793H95.7695V4.36292H107.851L107.853 8.54522Z" fill="white"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M126.155 25.2794H131.325L122.803 4.36145H118.348L109.823 25.2794H114.995L116.556 21.0955H124.592L126.155 25.2794ZM120.548 10.8281H120.599L122.954 16.9103H118.192L120.548 10.8281Z" fill="white"/>
<path fill-rule="evenodd" clip-rule="evenodd" d="M150 14.8082C150 8.72445 145.137 4.36292 139.352 4.36292H133.235V25.2793H139.352C145.162 25.2793 150 20.9192 150 14.8082ZM144.803 14.8335C144.777 18.3324 142.346 21.0198 138.788 21.0198H138.046V8.62164H138.788C142.371 8.62164 144.803 11.3347 144.803 14.8335Z" fill="white"/>
<path d="M38.4037 15.2638L28.8771 5.73713L19.3504 15.2638L16.7622 12.6756L28.8771 0.561646L40.991 12.6756L38.4037 15.2638Z" fill="white"/>
<path d="M38.4037 21.8914L28.8771 12.3647L19.3504 21.8914L16.7622 19.3032L28.8771 7.18921L40.991 19.3032L38.4037 21.8914Z" fill="white"/>
<path d="M40.991 25.9307L40.2975 26.6241L35.7826 25.8968L28.8771 18.9913L21.9715 25.8968L17.4566 26.6241L16.7622 25.9307L28.8771 13.8158L40.991 25.9307Z" fill="white"/>
<path d="M34.0509 25.6177L28.8772 24.7833L23.7026 25.6177L28.8772 20.4431L34.0509 25.6177Z" fill="white"/>
<path d="M16.0143 13.423L13.4272 16.01L15.9936 18.5763L18.5806 15.9893L16.0143 13.423Z" fill="white"/>
<path d="M12.7003 16.7372L10.1133 19.3242L15.9931 25.2041L18.5802 22.617L12.7003 16.7372Z" fill="white"/>
<path d="M16.1692 26.8316L11.7128 27.5498L6.80078 22.6378L9.38898 20.0505L16.1692 26.8316Z" fill="white"/>
<path d="M10.4626 27.752L0 29.4382L6.07481 23.3643L10.4626 27.752Z" fill="white"/>
<path d="M41.7406 13.4231L39.1738 15.989L41.7611 18.5771L44.3278 16.0113L41.7406 13.4231Z" fill="white"/>
<path d="M45.0522 16.7337L39.1724 22.6135L41.76 25.2012L47.6399 19.3214L45.0522 16.7337Z" fill="white"/>
<path d="M50.9548 22.6387L46.0428 27.5498L41.5854 26.8316L48.3665 20.0505L50.9548 22.6387Z" fill="white"/>
<path d="M57.7542 29.4382L47.2925 27.752L51.6803 23.3643L57.7542 29.4382Z" fill="white"/></svg>'''

def num(v):
    if v in (None, ""):
        return None
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except ValueError:
        return None

def load(name):
    p = os.path.join(VOL_DIR, name)
    if not os.path.exists(p):
        return []
    with open(p, newline="") as f:
        return list(csv.DictReader(f))

def series(rows):
    """Shape one granularity's rows into the JSON the dashboard charts consume."""
    cols = ["ic_success_email", "ic_support_email", "ic_chat", "ic_total",
            "lin_cus_created", "lin_cus_p0", "lin_cus_p1",
            "lin_rep_created", "lin_rep_p0", "lin_rep_p1", "lin_total_created",
            "grand_total", "metric_created_completed_1wk_pct",
            "ic_avg_first_response_min", "ic_avg_resolution_hr",
            "support_marco_escalations"]
    out = {"labels": [r["period_label"] for r in rows]}
    for c in cols:
        out[c] = [num(r.get(c, "")) for r in rows]
    return out

def kpi(rows):
    if not rows:
        return {}
    last = rows[-1]
    return {k: num(last.get(k, "")) for k in
            ["ic_total", "ic_support_email", "ic_success_email", "ic_chat",
             "lin_total_created", "grand_total", "metric_created_completed_1wk_pct"]}

def main():
    data = {g: series(load(f"volume_{g}.csv")) for g in ("daily", "weekly", "monthly")}
    kpis = kpi(load("volume_weekly.csv")) or kpi(load("volume_daily.csv"))
    html = TEMPLATE.replace("/*DATA*/", json.dumps(data)) \
                   .replace("/*KPI*/", json.dumps(kpis)) \
                   .replace("__WORDMARK__", TREAD_WORDMARK) \
                   .replace("__DATE__", TODAY)
    os.makedirs(VOL_DIR, exist_ok=True)
    with open(os.path.join(VOL_DIR, "index.html"), "w") as f:
        f.write(html)
    print(f"wrote {VOL_DIR}/index.html")

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Volume Pulse · Tread Customer Ops</title>
<link href="https://fonts.googleapis.com/css2?family=Golos+Text:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{--dark:#132732;--dark2:#0E1F2A;--panel:#17303D;--line:#23414F;--ink:#EAF2F6;--mut:#8FA8B4;--yellow:#FFE500;--amber:#FFAA13;--chat:#58C7C2;--cus:#9B8CFF;--rep:#4DA3FF;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Golos Text',system-ui,sans-serif;background:var(--dark2);color:var(--ink)}
nav{display:flex;align-items:center;justify-content:space-between;padding:16px 28px;background:var(--dark);border-bottom:1px solid var(--line)}
.nav-logo{display:flex;align-items:center;gap:14px;text-decoration:none}
.nav-divider{width:1px;height:22px;background:var(--line)}
.nav-title{font-size:.85rem;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--mut)}
.nav-back{color:var(--mut);text-decoration:none;font-size:.85rem;font-weight:600}
.nav-back:hover{color:var(--ink)}
.hero{padding:34px 28px 8px;max-width:1180px;margin:0 auto}
.hero h1{font-size:1.9rem;font-weight:800;display:flex;align-items:center;gap:12px}
.hero h1 .dot{width:11px;height:11px;border-radius:50%;background:var(--yellow);box-shadow:0 0 12px var(--yellow)}
.hero p{color:var(--mut);max-width:760px;margin-top:8px;font-size:.95rem;line-height:1.5}
.wrap{max-width:1180px;margin:0 auto;padding:18px 28px 60px}
.toggle{display:inline-flex;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:4px;margin:18px 0 22px}
.toggle button{font-family:inherit;font-weight:700;font-size:.82rem;color:var(--mut);background:none;border:none;padding:8px 18px;border-radius:7px;cursor:pointer}
.toggle button.on{background:var(--yellow);color:#10222C}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:8px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:16px 18px}
.kpi .l{font-size:.72rem;text-transform:uppercase;letter-spacing:.06em;color:var(--mut);margin-bottom:6px}
.kpi .v{font-size:1.9rem;font-weight:800;line-height:1}
.kpi .s{font-size:.74rem;color:var(--mut);margin-top:5px}
.kpi .v.y{color:var(--yellow)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-top:18px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:20px}
.card.full{grid-column:1/-1}
.card h2{font-size:.95rem;font-weight:700;margin-bottom:2px}
.card .meta{font-size:.78rem;color:var(--mut);margin-bottom:14px}
canvas{max-height:300px}
.note{font-size:.78rem;color:var(--mut);margin-top:10px;border-left:2px solid var(--amber);padding-left:10px}
.dl{display:flex;gap:10px;flex-wrap:wrap;margin-top:24px}
.dl a{font-size:.78rem;font-weight:700;color:#10222C;background:var(--yellow);padding:8px 14px;border-radius:8px;text-decoration:none}
.dl a.alt{background:transparent;color:var(--mut);border:1px solid var(--line)}
.foot{color:var(--mut);font-size:.76rem;margin-top:30px;border-top:1px solid var(--line);padding-top:16px;line-height:1.6}
@media(max-width:860px){.grid{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}}
</style></head><body>
<nav>
  <a class="nav-logo" href="../">__WORDMARK__<div class="nav-divider"></div><span class="nav-title">Volume Pulse</span></a>
  <a class="nav-back" href="../">← Hub</a>
</nav>
<div class="hero">
  <h1><span class="dot"></span>Customer Ops Volume Pulse</h1>
  <p>The source of truth for support volume trends — Intercom (success@, support@, chat) and Linear (CUS &amp; REP issues), tracked daily, weekly, and monthly. Refreshed every morning.</p>
</div>
<div class="wrap">
  <div class="toggle" id="tg">
    <button data-g="daily">Daily</button>
    <button data-g="weekly" class="on">Weekly</button>
    <button data-g="monthly">Monthly</button>
  </div>
  <div class="kpis" id="kpis"></div>
  <div class="grid">
    <div class="card full"><h2>Volume by channel</h2><div class="meta">Intercom success@ · support@ · chat, stacked, with Linear issues created</div><canvas id="cVol"></canvas></div>
    <div class="card"><h2>Linear issues created — CUS vs REP</h2><div class="meta">with P0 / P1 priority overlay</div><canvas id="cLin"></canvas></div>
    <div class="card"><h2>1 Metric — same-window completion</h2><div class="meta">% of CUS+REP issues created &amp; completed within 7 days</div><canvas id="cMetric"></canvas>
      <div class="note">The most recent period is partial — issues created in the last 7 days may not have had time to complete, so the rate is understated until the window closes.</div></div>
    <div class="card full"><h2>Intercom response &amp; handle time</h2><div class="meta">avg first admin reply (min) · avg time to resolution (hrs)</div><canvas id="cTime"></canvas></div>
  </div>
  <div class="dl">
    <a href="volume_daily.csv" download>↓ Daily CSV</a>
    <a href="volume_weekly.csv" download>↓ Weekly CSV</a>
    <a href="volume_monthly.csv" download>↓ Monthly CSV</a>
    <a class="alt" href="../team-throughput/">Team Throughput →</a>
  </div>
  <div class="foot">
    Sources: Intercom Conversations API (channel = source type; success@ via the “Success@ Email” tag) · Linear GraphQL (CUS &amp; REP, P0/P1 via the Customer Prioritization label group). <br>
    <strong>support_marco_escalations</strong> is a manual column (prod/eng escalations, found via Slack) — preserved across refreshes and editable directly in the CSV. Phone (Quo) is not yet included. Last refreshed __DATE__.
  </div>
</div>
<script>
const DATA=/*DATA*/, KPI=/*KPI*/;
const C={ink:'#EAF2F6',mut:'#8FA8B4',grid:'#23414F',yellow:'#FFE500',amber:'#FFAA13',chat:'#58C7C2',cus:'#9B8CFF',rep:'#4DA3FF'};
Chart.defaults.color=C.mut;Chart.defaults.borderColor=C.grid;Chart.defaults.font.family="'Golos Text',sans-serif";
let g='weekly',charts={};
function ds(label,key,color,extra){return Object.assign({label,data:DATA[g][key],borderColor:color,backgroundColor:color},extra||{});}
function mkKpis(){
  const k=KPI||{};
  const items=[
    ['Intercom total','ic_total',k.ic_total,'support+success+chat'],
    ['support@','ic_support_email',k.ic_support_email,'latest period'],
    ['Linear created','lin_total_created',k.lin_total_created,'CUS + REP'],
    ['1 Metric','metric_created_completed_1wk_pct',(k.metric_created_completed_1wk_pct==null?'—':k.metric_created_completed_1wk_pct+'%'),'same-window completion',true],
  ];
  document.getElementById('kpis').innerHTML=items.map(i=>
    `<div class="kpi"><div class="l">${i[0]}</div><div class="v ${i[4]?'y':''}">${i[2]==null?'—':i[2]}</div><div class="s">${i[3]}</div></div>`).join('');
}
function draw(){
  Object.values(charts).forEach(c=>c.destroy());charts={};
  const L=DATA[g].labels;
  const stack={stacked:true};
  charts.v=new Chart(cVol,{type:'bar',data:{labels:L,datasets:[
    ds('success@','ic_success_email',C.yellow,{stack:'ic'}),
    ds('support@','ic_support_email',C.amber,{stack:'ic'}),
    ds('chat','ic_chat',C.chat,{stack:'ic'}),
    ds('Linear created','lin_total_created',C.rep,{stack:'lin',type:'line',fill:false,tension:.3,borderWidth:2,pointRadius:2}),
  ]},options:{responsive:true,scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,beginAtZero:true}},plugins:{legend:{labels:{boxWidth:12}}}}});
  charts.l=new Chart(cLin,{type:'bar',data:{labels:L,datasets:[
    ds('CUS created','lin_cus_created',C.cus),ds('REP created','lin_rep_created',C.rep),
    ds('P0','lin_cus_p0','#FF5C5C',{type:'line',borderWidth:0,pointRadius:0,hidden:false,stack:undefined}),
  ].slice(0,2).concat([
    {label:'P0 (CUS+REP)',data:L.map((_,i)=>(DATA[g].lin_cus_p0[i]||0)+(DATA[g].lin_rep_p0[i]||0)),type:'line',borderColor:'#FF5C5C',backgroundColor:'#FF5C5C',tension:.3,borderWidth:2,pointRadius:2},
    {label:'P1 (CUS+REP)',data:L.map((_,i)=>(DATA[g].lin_cus_p1[i]||0)+(DATA[g].lin_rep_p1[i]||0)),type:'line',borderColor:C.amber,backgroundColor:C.amber,tension:.3,borderWidth:2,pointRadius:2},
  ])},options:{responsive:true,scales:{x:{grid:{display:false}},y:{beginAtZero:true}},plugins:{legend:{labels:{boxWidth:12}}}}});
  charts.m=new Chart(cMetric,{type:'line',data:{labels:L,datasets:[
    ds('% created & completed ≤7d','metric_created_completed_1wk_pct',C.yellow,{tension:.3,fill:true,backgroundColor:'rgba(255,229,0,.08)',borderWidth:2,pointRadius:2})
  ]},options:{responsive:true,scales:{x:{grid:{display:false}},y:{beginAtZero:true,max:100,ticks:{callback:v=>v+'%'}}},plugins:{legend:{display:false}}}});
  charts.t=new Chart(cTime,{data:{labels:L,datasets:[
    ds('avg first reply (min)','ic_avg_first_response_min',C.chat,{type:'line',tension:.3,borderWidth:2,pointRadius:2,yAxisID:'y'}),
    ds('avg resolution (hrs)','ic_avg_resolution_hr',C.amber,{type:'line',tension:.3,borderWidth:2,pointRadius:2,yAxisID:'y1'}),
  ]},options:{responsive:true,scales:{x:{grid:{display:false}},y:{position:'left',beginAtZero:true,title:{display:true,text:'min'}},y1:{position:'right',beginAtZero:true,grid:{drawOnChartArea:false},title:{display:true,text:'hrs'}}},plugins:{legend:{labels:{boxWidth:12}}}}});
}
document.getElementById('tg').addEventListener('click',e=>{
  if(!e.target.dataset.g)return;
  g=e.target.dataset.g;
  [...tg.children].forEach(b=>b.classList.toggle('on',b.dataset.g===g));
  draw();
});
mkKpis();draw();
</script>
</body></html>"""

if __name__ == "__main__":
    main()
