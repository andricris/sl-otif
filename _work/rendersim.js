const path = require("path"), fs = require("fs");
const ExcelJS = require(path.join("C:", "Users", "LOGISTIK1", "Downloads", "sl otif", "web", "exceljs.min.js"));
const core = require(path.join("C:", "Users", "LOGISTIK1", "Downloads", "sl otif", "web", "sl-otif-core.js"));
const base = path.join("C:", "Users", "LOGISTIK1", "Downloads", "sl otif");

function tone(v){ return v>=0.99?"var(--good)":v>=0.9?"var(--warn)":"var(--accent)"; }
function pct(v){ return (v*100).toFixed(2)+"%"; }
function stat(v){ return '<span class="stat"><span class="dot" style="background:'+tone(v)+'"></span>'+pct(v)+"</span>"; }

(async () => {
  for (const f of ["us.xlsx","heonz.xlsx","muncul.xlsx"]) {
    const p = path.join(base, f);
    if (!fs.existsSync(p)) { console.log(f, "MISSING"); continue; }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(fs.readFileSync(p));
    const st = core.build(wb, 7);
    console.log("=" .repeat(70));
    console.log(f, "order=", st.order.join(","));
    const agg = st.agg;
    st.order.forEach(m => {
      const d = agg[m], f2 = d.lateL + d.ontL, gg = d.lateS + d.ontS;
      console.log(`  ${m}: lateL=${d.lateL} lateS=${d.lateS} ontL=${d.ontL} ontS=${d.ontS}  f=${f2} gg=${gg} -> FULFIL=${f2?pct(gg/f2):"—"} OTIF=${f2?pct(d.ontS/f2):"—"}`);
    });
    const g = st.grand, F = g.lateL + g.ontL, G = g.lateS + g.ontS;
    console.log(`  TOTAL: F=${F} G=${G} FULFIL=${F?pct(G/F):"—"} OTIF=${F?pct(g.ontS/F):"—"}`);
  }
})().catch(e => { console.error(e); process.exit(1); });
