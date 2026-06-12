#!/usr/bin/env python3
"""
Render the Support Pulse dashboard (customer-pulse/support-pulse/index.html) from
the three CSVs written by refresh_volume.py. Tread-branded, Chart.js, with a
Daily / Weekly / Monthly toggle. One idea per chart, built for at-a-glance reading.
No external data calls — pure CSV → HTML.
"""
import csv, datetime, json, os

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
VOL_DIR = os.path.join(OUT_DIR, "support-pulse")
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
    out = {"labels": [r["period_label"] for r in rows],
           "period_start": [r.get("period_start", "") for r in rows],
           "period_end": [r.get("period_end", "") for r in rows]}
    for c in cols:
        out[c] = [num(r.get(c, "")) for r in rows]
    return out

def load_categories():
    """Long-format categories CSV → {date: {main: {sub: count}}}."""
    out = {}
    p = os.path.join(VOL_DIR, "support_categories_daily.csv")
    if not os.path.exists(p):
        return out
    with open(p, newline="") as f:
        for r in csv.DictReader(f):
            try:
                c = int(float(r["count"]))
            except (ValueError, KeyError):
                continue
            out.setdefault(r["date"], {}).setdefault(r["main"], {})[r["sub"]] = c
    return out

def load_usertypes():
    """Long-format usertypes CSV → {date: {usertype: count}}."""
    out = {}
    p = os.path.join(VOL_DIR, "support_usertypes_daily.csv")
    if not os.path.exists(p):
        return out
    with open(p, newline="") as f:
        for r in csv.DictReader(f):
            try:
                out.setdefault(r["date"], {})[r["usertype"]] = int(float(r["count"]))
            except (ValueError, KeyError):
                continue
    return out

def main():
    data = {g: series(load(f"support_{g}.csv")) for g in ("daily", "weekly", "monthly")}
    cats = load_categories()
    utypes = load_usertypes()
    html = (TEMPLATE.replace("/*DATA*/", json.dumps(data))
                    .replace("/*CATS*/", json.dumps(cats))
                    .replace("/*UTYPES*/", json.dumps(utypes))
                    .replace("__WORDMARK__", TREAD_WORDMARK)
                    .replace("__DATE__", TODAY))
    os.makedirs(VOL_DIR, exist_ok=True)
    with open(os.path.join(VOL_DIR, "index.html"), "w") as f:
        f.write(html)
    print(f"wrote {VOL_DIR}/index.html")

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Support Pulse · Tread Customer Ops</title>
<link href="https://fonts.googleapis.com/css2?family=Golos+Text:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{--dark:#132732;--dark2:#0E1F2A;--panel:#17303D;--line:#23414F;--ink:#EAF2F6;--mut:#8FA8B4;--yellow:#FFE500;--amber:#FFAA13;--chat:#58C7C2;--cus:#9B8CFF;--rep:#4DA3FF;--red:#FF6B6B;--green:#4ADE80;}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Golos Text',system-ui,sans-serif;background:var(--dark2);color:var(--ink)}
nav{display:flex;align-items:center;justify-content:space-between;padding:16px 28px;background:var(--dark);border-bottom:1px solid var(--line)}
.nav-logo{display:flex;align-items:center;gap:14px;text-decoration:none}
.nav-divider{width:1px;height:22px;background:var(--line)}
.nav-title{font-size:.85rem;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--mut)}
.nav-back{color:var(--mut);text-decoration:none;font-size:.85rem;font-weight:600}
.nav-back:hover{color:var(--ink)}
.hero{padding:32px 28px 4px;max-width:1180px;margin:0 auto}
.hero h1{font-size:1.9rem;font-weight:800;display:flex;align-items:center;gap:12px}
.hero h1 .dot{width:11px;height:11px;border-radius:50%;background:var(--yellow);box-shadow:0 0 12px var(--yellow)}
.hero p{color:var(--mut);max-width:780px;margin-top:8px;font-size:.95rem;line-height:1.5}
.wrap{max-width:1180px;margin:0 auto;padding:14px 28px 60px}
.vtabs{display:flex;gap:4px;border-bottom:1px solid var(--line);margin:18px 0 4px}
.vtabs button{font-family:inherit;font-weight:700;font-size:.9rem;color:var(--mut);background:none;border:none;border-bottom:2px solid transparent;padding:10px 16px;cursor:pointer;margin-bottom:-1px}
.vtabs button.on{color:var(--ink);border-bottom-color:var(--yellow)}
.bar{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin:16px 0 20px}
.tbl{width:100%;border-collapse:collapse;font-size:.82rem}
.tbl th{text-align:left;color:var(--mut);font-weight:600;text-transform:uppercase;letter-spacing:.05em;font-size:.68rem;padding:8px 10px;border-bottom:1px solid var(--line)}
.tbl td{padding:7px 10px;border-bottom:1px solid rgba(35,65,79,.5)}
.tbl td.n{text-align:right;font-variant-numeric:tabular-nums;font-weight:700}
.tbl td.bar-cell{width:34%}
.tbl .minib{height:8px;border-radius:4px;background:var(--cus)}
.tbl tr.maincat td{font-weight:700;color:var(--ink);background:rgba(255,229,0,.05)}
.tbl tr.subcat td:first-child{padding-left:24px;color:var(--mut)}
.toggle{display:inline-flex;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:4px}
.toggle button{font-family:inherit;font-weight:700;font-size:.82rem;color:var(--mut);background:none;border:none;padding:8px 18px;border-radius:7px;cursor:pointer}
.toggle button.on{background:var(--yellow);color:#10222C}
.asof{font-size:.76rem;color:var(--mut)}
.range{display:flex;align-items:center;gap:10px;font-size:.76rem;color:var(--mut)}
.range label{display:flex;align-items:center;gap:6px}
.range select{font-family:inherit;font-size:.8rem;color:var(--ink);background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:6px 8px;max-width:150px}
.range #resetRange{font-family:inherit;font-size:.74rem;font-weight:600;color:var(--mut);background:none;border:1px solid var(--line);border-radius:8px;padding:6px 10px;cursor:pointer}
.range #resetRange:hover{color:var(--ink);border-color:var(--mut)}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:16px 18px}
.kpi .l{font-size:.72rem;text-transform:uppercase;letter-spacing:.06em;color:var(--mut);margin-bottom:8px}
.kpi .v{font-size:2rem;font-weight:800;line-height:1}
.kpi .v.y{color:var(--yellow)}
.kpi .d{font-size:.76rem;margin-top:7px;font-weight:600}
.kpi .d.up{color:var(--green)}.kpi .d.down{color:var(--red)}.kpi .d.flat{color:var(--mut)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:20px 20px 14px}
.card.full{grid-column:1/-1}
.card h2{font-size:1rem;font-weight:700;margin-bottom:2px}
.card .meta{font-size:.78rem;color:var(--mut);margin-bottom:16px}
.card .wrapc{position:relative;height:300px}
.card.sm .wrapc{height:230px}
.note{font-size:.76rem;color:var(--mut);margin-top:12px;border-left:2px solid var(--amber);padding-left:10px}
.dl{display:flex;gap:10px;flex-wrap:wrap;margin-top:24px}
.dl a{font-size:.78rem;font-weight:700;color:#10222C;background:var(--yellow);padding:9px 15px;border-radius:8px;text-decoration:none}
.dl a.alt{background:transparent;color:var(--mut);border:1px solid var(--line)}
.foot{color:var(--mut);font-size:.76rem;margin-top:28px;border-top:1px solid var(--line);padding-top:16px;line-height:1.6}
@media(max-width:860px){.grid{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}}
</style></head><body>
<nav>
  <a class="nav-logo" href="../">__WORDMARK__<div class="nav-divider"></div><span class="nav-title">Support Pulse</span></a>
  <a class="nav-back" href="../">← Hub</a>
</nav>
<div class="hero">
  <h1><span class="dot"></span>Support Pulse</h1>
  <p>The source of truth for support volume trends — Intercom (success@, support@, chat) and Linear (CUS &amp; REP issues), tracked daily, weekly, and monthly. Refreshed every morning.</p>
</div>
<div class="wrap">
  <div class="bar">
    <div class="toggle" id="tg">
      <button data-g="daily">Daily</button>
      <button data-g="weekly" class="on">Weekly</button>
      <button data-g="monthly">Monthly</button>
    </div>
    <div class="range">
      <label>From <select id="from"></select></label>
      <label>To <select id="to"></select></label>
      <button id="resetRange" type="button">Reset</button>
    </div>
    <div class="asof" id="asof"></div>
  </div>

  <div class="vtabs" id="vt">
    <button data-v="overview" class="on">Overview</button>
    <button data-v="type">By type</button>
    <button data-v="usertype">By user type</button>
  </div>

 <div id="viewOverview">
  <div class="kpis" id="kpis"></div>

  <div class="grid">
    <div class="card full">
      <h2>Support volume by channel</h2><div class="meta">Inbound conversations — success@ · support@ · chat, stacked</div>
      <div class="wrapc"><canvas id="cVol"></canvas></div>
    </div>

    <div class="card full">
      <h2>The 1 Metric — same-window completion</h2><div class="meta">% of CUS+REP issues created &amp; completed within 7 days</div>
      <div class="wrapc"><canvas id="cMetric"></canvas></div>
      <div class="note">The most recent period is partial — issues created in the last 7 days may not have had time to complete, so the rate reads low until the window closes.</div>
    </div>

    <div class="card sm"><h2>Linear issues created</h2><div class="meta">CUS vs REP</div><div class="wrapc"><canvas id="cLin"></canvas></div></div>
    <div class="card sm"><h2>Priority pressure — P0 &amp; P1</h2><div class="meta">CUS + REP issues labeled P0 / P1</div><div class="wrapc"><canvas id="cPrio"></canvas></div></div>

    <div class="card sm"><h2>Avg first response</h2><div class="meta">Intercom — minutes to first admin reply</div><div class="wrapc"><canvas id="cResp"></canvas></div></div>
    <div class="card sm"><h2>Avg resolution time</h2><div class="meta">Intercom — hours to close</div><div class="wrapc"><canvas id="cRes"></canvas></div></div>
  </div>
 </div>

 <div id="viewType" style="display:none">
  <div class="grid">
    <div class="card full">
      <h2>Volume by type — trend</h2><div class="meta">Inbound conversations by main category over the selected range</div>
      <div class="wrapc"><canvas id="cTypeTrend"></canvas></div>
    </div>
    <div class="card full">
      <h2>Volume by type — totals</h2><div class="meta">Main categories segmented by sub-category, over the selected range</div>
      <div class="wrapc"><canvas id="cTypeBar"></canvas></div>
    </div>
    <div class="card full">
      <h2>Category breakdown</h2><div class="meta">Main category › sub-category, by conversation volume</div>
      <table class="tbl"><thead><tr><th>Category</th><th style="text-align:right">Volume</th><th>Share</th></tr></thead><tbody id="catTbody"></tbody></table>
      <div class="note">Type is derived from Intercom conversation tags + subject/body text (one category per conversation). “Uncategorized” = no recognizable topic. Mapping lives in <code>refresh_volume.py</code> and is easy to refine.</div>
    </div>
  </div>
 </div>

 <div id="viewUtype" style="display:none">
  <div class="grid">
    <div class="card full">
      <h2>Which user types need the most support — trend</h2><div class="meta">Conversations by requester persona over the selected range</div>
      <div class="wrapc"><canvas id="cUtTrend"></canvas></div>
    </div>
    <div class="card full">
      <h2>Support volume by user type</h2><div class="meta">Total conversations per persona, over the selected range</div>
      <div class="wrapc"><canvas id="cUtBar"></canvas></div>
    </div>
    <div class="card full">
      <h2>User-type breakdown</h2><div class="meta">Requester persona by conversation volume</div>
      <table class="tbl"><thead><tr><th>User type</th><th style="text-align:right">Volume</th><th>Share</th></tr></thead><tbody id="utTbody"></tbody></table>
      <div class="note">Persona is assigned per conversation: first from what the request is about (content), then the requester’s Intercom platform <em>Roles</em> attribute, else “Unknown.” A user can hold several platform roles; this picks one primary persona so the split sums to total volume.</div>
    </div>
  </div>
 </div>

  <div class="dl">
    <a href="support_pulse.xlsx" download>↓ Support Pulse workbook (.xlsx — Daily · Weekly · Monthly tabs)</a>
    <a class="alt" href="../team-throughput/">Team Throughput →</a>
  </div>
  <div class="foot">
    Sources: Intercom Conversations API (channel = source type; success@ via the “Success@ Email” tag) · Linear GraphQL (CUS &amp; REP; P0/P1 via the Customer Prioritization label group). <br>
    <strong>Marco prod/eng escalations</strong> is a manual column (found via Slack) — type counts into the workbook (e.g. the Weekly tab); the daily refresh reads them back and never overwrites them. Phone (Quo) is not yet included. Last refreshed __DATE__.
  </div>
</div>
<script>
const DATA=/*DATA*/;
const CATS=/*CATS*/;
const UTYPES=/*UTYPES*/;
const UTCOLORS={'Driver':'#4DA3FF','Dispatcher':'#58C7C2','Biller':'#FFAA13','Reporting':'#FFE500','Foreman':'#A3E635','Manager':'#9B8CFF','Company Admin':'#FF8FB1','IT Admin':'#F472B6','Platform Admin':'#C084FC','Viewer':'#7DD3FC','Unknown':'#6B828C'};
const utColor=(u,i)=>UTCOLORS[u]||['#4DA3FF','#58C7C2','#FFAA13','#FFE500','#A3E635','#9B8CFF','#FF8FB1'][i%7];
const C={ink:'#EAF2F6',mut:'#8FA8B4',grid:'rgba(143,168,180,.14)',yellow:'#FFE500',amber:'#FFAA13',chat:'#58C7C2',cus:'#9B8CFF',rep:'#4DA3FF',red:'#FF6B6B',green:'#4ADE80'};
const MAINCOLORS={'Account & Access':'#9B8CFF','Tickets':'#4DA3FF','Dispatch':'#58C7C2','Jobs':'#4ADE80','Billing & Rates':'#FFAA13','Integrations':'#FF8FB1','Reporting & Insights':'#FFE500','Mobile & App':'#7DD3FC','Compliance':'#F472B6','Feature Requests':'#A3E635','Bugs':'#FF6B6B','Other':'#6B828C'};
const catColor=(m,i)=>MAINCOLORS[m]||['#9B8CFF','#4DA3FF','#58C7C2','#4ADE80','#FFAA13','#FF8FB1','#FFE500'][i%7];
Chart.defaults.color=C.mut;Chart.defaults.borderColor=C.grid;Chart.defaults.font.family="'Golos Text',sans-serif";Chart.defaults.font.size=11;
Chart.defaults.plugins.legend.labels.boxWidth=10;Chart.defaults.plugins.legend.labels.usePointStyle=true;Chart.defaults.plugins.legend.position='bottom';
let g='weekly', charts={}, range={from:0,to:0}, cur=null;
const X={grid:{display:false},ticks:{maxRotation:0,autoSkip:true,maxTicksLimit:9}};
const Y=(opts={})=>Object.assign({beginAtZero:true,grid:{color:C.grid,drawTicks:false},border:{display:false}},opts);
const last=a=>{for(let i=a.length-1;i>=0;i--)if(a[i]!=null)return a[i];return null;};
const prev=a=>{let seen=0;for(let i=a.length-1;i>=0;i--){if(a[i]!=null){seen++;if(seen===2)return a[i];}}return null;};
const MON=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function ymd(s){const p=String(s).split('-').map(Number);return{y:p[0],m:p[1],d:p[2]};}
// Human date labels — weeks/months shown as date ranges, not week numbers.
function disp(gr,i){
  const s=DATA[gr].period_start[i], e=DATA[gr].period_end[i];
  if(!s)return DATA[gr].labels[i];
  const a=ymd(s);
  if(gr==='daily')  return `${MON[a.m-1]} ${a.d}`;
  if(gr==='monthly')return `${MON[a.m-1]} ${a.y}`;
  const b=ymd(e);
  return a.m===b.m ? `${MON[a.m-1]} ${a.d}–${b.d}` : `${MON[a.m-1]} ${a.d} – ${MON[b.m-1]} ${b.d}`;
}
function allDisp(){return DATA[g].labels.map((_,i)=>disp(g,i));}
function populateRange(reset){
  const d=allDisp(), n=d.length;
  if(reset){range.from=0;range.to=n-1;}
  range.from=Math.min(range.from,n-1); range.to=Math.min(range.to,n-1);
  const fill=(el,sel)=>{el.innerHTML=d.map((x,i)=>`<option value="${i}"${i===sel?' selected':''}>${x}</option>`).join('');};
  fill(document.getElementById('from'),range.from);
  fill(document.getElementById('to'),range.to);
}
function view(){
  const f=Math.min(range.from,range.to), t=Math.max(range.from,range.to);
  const src=DATA[g], o={labels:allDisp().slice(f,t+1)};
  for(const k in src){if(k!=='labels' && Array.isArray(src[k]))o[k]=src[k].slice(f,t+1);}
  return o;
}
function fmtDelta(c,p,opts){
  if(c==null||p==null)return {cls:'flat',txt:'—'};
  const d=c-p, pts=opts&&opts.pts;
  const good = opts&&opts.higherBetter ? d>0 : null;
  const cls = d===0?'flat':(good===null?'flat':(good?'up':'down'));
  const arrow = d>0?'▲':(d<0?'▼':'■');
  const val = pts ? Math.abs(d).toFixed(1)+' pts' : (Math.abs(d)%1?Math.abs(d).toFixed(1):Math.abs(d));
  return {cls, txt:`${arrow} ${val} vs prior`};
}
function mkKpis(){
  const defs=[
    {l:'Support volume',key:'ic_total'},
    {l:'support@ inbound',key:'ic_support_email'},
    {l:'Linear created',key:'lin_total_created'},
    {l:'1 Metric',key:'metric_created_completed_1wk_pct',pct:true,pts:true,higherBetter:true},
  ];
  document.getElementById('kpis').innerHTML=defs.map(o=>{
    const c=last(cur[o.key]), p=prev(cur[o.key]);
    const dd=fmtDelta(c,p,{pts:o.pts,higherBetter:o.higherBetter});
    const v=c==null?'—':(o.pct?c+'%':c);
    return `<div class="kpi"><div class="l">${o.l}</div><div class="v ${o.pct?'y':''}">${v}</div><div class="d ${dd.cls}">${dd.txt}</div></div>`;
  }).join('');
}
function area(key,color){return {data:cur[key],borderColor:color,backgroundColor:color+'33',fill:true,stack:'s',tension:.35,borderWidth:2,pointRadius:0,pointHoverRadius:4};}
function line(key,color,fill){return {data:cur[key],borderColor:color,backgroundColor:fill?color+'22':color,fill:!!fill,tension:.35,borderWidth:2.5,pointRadius:0,pointHoverRadius:4};}
function sum2(a,b){return cur[a].map((v,i)=>(v||0)+(cur[b][i]||0));}

function drawOverview(){
  const L=cur.labels;
  charts.v=new Chart(cVol,{type:'line',data:{labels:L,datasets:[
    Object.assign({label:'success@'},area('ic_success_email',C.yellow)),
    Object.assign({label:'support@'},area('ic_support_email',C.amber)),
    Object.assign({label:'chat'},area('ic_chat',C.chat)),
  ]},options:{maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{x:X,y:Y({stacked:true})}}});

  charts.m=new Chart(cMetric,{type:'line',data:{labels:L,datasets:[
    Object.assign({label:'% created & completed ≤7d'},line('metric_created_completed_1wk_pct',C.yellow,true)),
  ]},options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:X,y:Y({max:100,ticks:{callback:v=>v+'%'}})}}});

  charts.l=new Chart(cLin,{type:'bar',data:{labels:L,datasets:[
    {label:'CUS',data:cur.lin_cus_created,backgroundColor:C.cus,borderRadius:4,maxBarThickness:22},
    {label:'REP',data:cur.lin_rep_created,backgroundColor:C.rep,borderRadius:4,maxBarThickness:22},
  ]},options:{maintainAspectRatio:false,scales:{x:X,y:Y()}}});

  charts.p=new Chart(cPrio,{type:'bar',data:{labels:L,datasets:[
    {label:'P0',data:sum2('lin_cus_p0','lin_rep_p0'),backgroundColor:C.red,borderRadius:4,maxBarThickness:22},
    {label:'P1',data:sum2('lin_cus_p1','lin_rep_p1'),backgroundColor:C.amber,borderRadius:4,maxBarThickness:22},
  ]},options:{maintainAspectRatio:false,scales:{x:X,y:Y({ticks:{precision:0}})}}});

  charts.r=new Chart(cResp,{type:'line',data:{labels:L,datasets:[
    Object.assign({label:'min'},line('ic_avg_first_response_min',C.chat,true)),
  ]},options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:X,y:Y()}}});

  charts.x=new Chart(cRes,{type:'line',data:{labels:L,datasets:[
    Object.assign({label:'hrs'},line('ic_avg_resolution_hr',C.amber,true)),
  ]},options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:X,y:Y()}}});
}
// ── Volume-by-type helpers (CATS is per-day; aggregate by date window) ──
let vmode='overview';
const winDates=()=>{const f=Math.min(range.from,range.to),t=Math.max(range.from,range.to);return [DATA[g].period_start[f],DATA[g].period_end[t]];};
function sumCats(d0,d1){const acc={};for(const dt in CATS){if(dt>=d0&&dt<=d1){const day=CATS[dt];for(const m in day){acc[m]=acc[m]||{};for(const s in day[m])acc[m][s]=(acc[m][s]||0)+day[m][s];}}}return acc;}
function mainTotal(o){const t={};for(const m in o)t[m]=Object.values(o[m]).reduce((a,b)=>a+b,0);return t;}
function drawType(){
  const f=Math.min(range.from,range.to), t=Math.max(range.from,range.to);
  const L=cur.labels;
  const agg=sumCats(...winDates()), tot=mainTotal(agg);
  const mains=Object.keys(agg).sort((a,b)=>tot[b]-tot[a]);

  // Trend: one stacked-area series per main category across the visible periods
  const trendDs=mains.map((m,i)=>{
    const data=L.map((_,j)=>{const ps=DATA[g].period_start[f+j],pe=DATA[g].period_end[f+j];const a=sumCats(ps,pe);return a[m]?Object.values(a[m]).reduce((x,y)=>x+y,0):0;});
    const col=catColor(m,i);
    return {label:m,data,borderColor:col,backgroundColor:col+'33',fill:true,stack:'s',tension:.35,borderWidth:1.5,pointRadius:0,pointHoverRadius:4};
  });
  charts.tt=new Chart(cTypeTrend,{type:'line',data:{labels:L,datasets:trendDs},options:{maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{x:X,y:Y({stacked:true})},plugins:{legend:{labels:{boxWidth:10}}}}});

  // Totals: bar per main category, stacked by sub-category (legend hidden; table has detail)
  const barDs=[];
  mains.forEach((m,i)=>{const subs=Object.keys(agg[m]).sort((a,b)=>agg[m][b]-agg[m][a]);
    subs.forEach((s,k)=>{const data=mains.map(mm=>mm===m?agg[m][s]:0);
      barDs.push({label:`${m} › ${s}`,data,backgroundColor:shade(catColor(m,i),k),borderRadius:3,maxBarThickness:46,stack:'x'});});});
  charts.tb=new Chart(cTypeBar,{type:'bar',data:{labels:mains,datasets:barDs},options:{maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.dataset.label.split(' › ')[1]+': '+c.parsed.y}}},scales:{x:Object.assign({stacked:true},X),y:Y({stacked:true,ticks:{precision:0}})}}});

  // Breakdown table
  const grand=mains.reduce((a,m)=>a+tot[m],0)||1;
  let html='';
  mains.forEach((m,i)=>{
    html+=`<tr class="maincat"><td>${m}</td><td class="n">${tot[m]}</td><td class="bar-cell"><div class="minib" style="width:${Math.round(100*tot[m]/grand)}%;background:${catColor(m,i)}"></div></td></tr>`;
    Object.entries(agg[m]).sort((a,b)=>b[1]-a[1]).forEach(([s,c])=>{
      html+=`<tr class="subcat"><td>${s}</td><td class="n">${c}</td><td>${(100*c/grand).toFixed(1)}%</td></tr>`;});
  });
  document.getElementById('catTbody').innerHTML=html || '<tr><td colspan="3" style="color:var(--mut)">No tagged conversations in range.</td></tr>';
}
function shade(hex,k){ // lighten a hex color by step k for sub-segment contrast
  const n=parseInt(hex.slice(1),16); let r=n>>16,gg=(n>>8)&255,b=n&255;
  const f=1-Math.min(k*0.13,0.6); const mix=v=>Math.round(v*f+255*(1-f));
  return `rgb(${mix(r)},${mix(gg)},${mix(b)})`;
}
// ── By user type (UTYPES is per-day {usertype:count}) ──
function sumUt(d0,d1){const acc={};for(const dt in UTYPES){if(dt>=d0&&dt<=d1){for(const u in UTYPES[dt])acc[u]=(acc[u]||0)+UTYPES[dt][u];}}return acc;}
function drawUtype(){
  const f=Math.min(range.from,range.to), t=Math.max(range.from,range.to);
  const L=cur.labels;
  const agg=sumUt(...winDates());
  const types=Object.keys(agg).sort((a,b)=>agg[b]-agg[a]);
  // Trend: stacked area, one series per user type across visible periods
  const ds=types.map((u,i)=>{
    const data=L.map((_,j)=>{const a=sumUt(DATA[g].period_start[f+j],DATA[g].period_end[f+j]);return a[u]||0;});
    const col=utColor(u,i);
    return {label:u,data,borderColor:col,backgroundColor:col+'33',fill:true,stack:'s',tension:.35,borderWidth:1.5,pointRadius:0,pointHoverRadius:4};
  });
  charts.ut=new Chart(cUtTrend,{type:'line',data:{labels:L,datasets:ds},options:{maintainAspectRatio:false,interaction:{mode:'index',intersect:false},scales:{x:X,y:Y({stacked:true})},plugins:{legend:{labels:{boxWidth:10}}}}});
  // Totals bar (one bar per user type, sorted)
  charts.ub=new Chart(cUtBar,{type:'bar',data:{labels:types,datasets:[{data:types.map(u=>agg[u]),backgroundColor:types.map((u,i)=>utColor(u,i)),borderRadius:4,maxBarThickness:40}]},options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:X,y:Y({ticks:{precision:0}})}}});
  // Table
  const grand=types.reduce((a,u)=>a+agg[u],0)||1;
  document.getElementById('utTbody').innerHTML = types.map((u,i)=>
    `<tr><td><span style="display:inline-block;width:9px;height:9px;border-radius:50%;background:${utColor(u,i)};margin-right:8px"></span>${u}</td><td class="n">${agg[u]}</td><td>${(100*agg[u]/grand).toFixed(1)}%</td></tr>`
  ).join('') || '<tr><td colspan="3" style="color:var(--mut)">No conversations in range.</td></tr>';
}
function render(){
  cur=view();
  Object.values(charts).forEach(c=>c.destroy());charts={};
  if(vmode==='overview'){mkKpis();drawOverview();}
  else if(vmode==='type'){drawType();}
  else {drawUtype();}
  const L=cur.labels;
  document.getElementById('asof').textContent = !L.length?'' : (L.length===1?`Showing ${L[0]}`:`Showing ${L[0]} → ${L[L.length-1]}`);
}
document.getElementById('vt').addEventListener('click',e=>{
  if(!e.target.dataset.v)return;
  vmode=e.target.dataset.v;
  [...vt.children].forEach(b=>b.classList.toggle('on',b.dataset.v===vmode));
  document.getElementById('viewOverview').style.display = vmode==='overview'?'':'none';
  document.getElementById('viewType').style.display = vmode==='type'?'':'none';
  document.getElementById('viewUtype').style.display = vmode==='usertype'?'':'none';
  render();
});
document.getElementById('tg').addEventListener('click',e=>{
  if(!e.target.dataset.g)return;
  g=e.target.dataset.g;
  [...tg.children].forEach(b=>b.classList.toggle('on',b.dataset.g===g));
  populateRange(true); render();
});
document.getElementById('from').addEventListener('change',e=>{range.from=+e.target.value; populateRange(false); render();});
document.getElementById('to').addEventListener('change',e=>{range.to=+e.target.value; populateRange(false); render();});
document.getElementById('resetRange').addEventListener('click',()=>{populateRange(true); render();});
populateRange(true); render();
</script>
</body></html>"""

if __name__ == "__main__":
    main()
