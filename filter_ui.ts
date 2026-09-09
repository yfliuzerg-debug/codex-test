export function renderFilterDashboard(r: any): string {
  const meta = new Map<string, any>();
  for (const s of r.so || []) {
    if (!meta.has(String(s.sku || ""))) {
      meta.set(String(s.sku || ""), {
        s1: String(s.s1 || ""),
        s2: String(s.s2 || ""),
        s3: String(s.s3 || ""),
        brand: String(s.brand || ""),
        productName: String(s.skuName || ""),
      });
    }
  }

  const rows = (r.exportRows || []).map((x: any) => {
    const m = meta.get(String(x.SKU || "")) || {};
    const customer = String(x["经销商/客户"] || "");
    const customerCode = String(x["客户编码"] || "");
    const activity = String(x["活动"] || "");
    const cp = String(x.CP || "");
    const sku = String(x.SKU || "");
    const productName = String(x["产品名称"] || m.productName || "");
    const brand = String(x["品牌"] || m.brand || "");
    const s1 = String(m.s1 || "");
    const s2 = String(m.s2 || "");
    const s3 = String(m.s3 || "");
    const productGroup = String(x["产品组"] || s3 || s2 || s1 || brand || "");
    const groupPath = [s3, s2, s1, brand].filter(Boolean).join(" › ");
    const month = String(x["月份"] || "");
    return {
      customer,
      customerCode,
      customerKey: customer + "\u241f" + customerCode,
      activity,
      cp,
      month,
      channel: String(x["渠道"] || ""),
      brand,
      sku,
      productName,
      productGroup,
      groupPath,
      balance: Number(x["CP余额"] || 0),
      so: Number(x.SO || 0),
      forecastSO: Number(x["预测SO"] || 0),
      rate: x["费率"] == null ? null : Number(x["费率"]),
      net: x["净承载"] == null ? null : Number(x["净承载"]),
      forecastCapacity: x["预测承载"] == null ? null : Number(x["预测承载"]),
      visibleCapacity: x["计入可见承载"] == null ? null : Number(x["计入可见承载"]),
    };
  });

  const safeRows = JSON.stringify(rows).replace(/</g, "\\u003c").replace(/-->/g, "--\\>");
  const currentMonth = JSON.stringify(String(r.currentMonth || ""));

  return `<style>
.filter-wrap{margin:14px 0}.filter-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:12px}.filter-head h3{margin:0}.filter-reset{border:1px solid #cfd8e6;background:#fff;color:#1769ff;border-radius:8px;padding:7px 11px;cursor:pointer}.filter-grid{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:10px}.filter-dd{position:relative;border:1px solid #dbe3ef;border-radius:10px;background:#fff}.filter-dd summary{list-style:none;cursor:pointer;padding:10px 12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.filter-dd summary::-webkit-details-marker{display:none}.filter-menu{position:absolute;z-index:20;top:calc(100% + 4px);left:0;right:0;max-height:300px;overflow:auto;background:#fff;border:1px solid #cfd8e6;border-radius:10px;box-shadow:0 12px 30px rgba(20,32,51,.12);padding:7px}.filter-option{display:flex;align-items:center;gap:8px;padding:7px 6px;border-radius:7px;cursor:pointer}.filter-option:hover{background:#f4f7fb}.filter-option span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.filter-note{margin-top:10px}.filter-kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin:14px 0}.filter-kpi{background:#fff;border:1px solid #dbe3ef;border-radius:12px;padding:12px}.filter-kpi span{display:block;color:#6b7688;font-size:12px}.filter-kpi b{display:block;font-size:21px;margin-top:2px}.filter-table-wrap{overflow:auto}.filter-table{min-width:1250px}.filter-empty{padding:24px;text-align:center;color:#6b7688}.filter-count{margin:4px 0 10px}.filter-summary-title{display:flex;align-items:baseline;justify-content:space-between;gap:12px}.filter-summary-title h3{margin:0}.filter-detail{margin-top:14px}.filter-detail table{min-width:1500px}@media(max-width:1000px){.filter-grid{grid-template-columns:repeat(2,minmax(150px,1fr))}.filter-kpis{grid-template-columns:repeat(3,1fr)}}@media(max-width:620px){.filter-grid{grid-template-columns:1fr}.filter-kpis{grid-template-columns:repeat(2,1fr)}}
</style>
<div class="p filter-wrap">
  <div class="filter-head"><div><h3>结果筛选</h3><div class="muted">一次计算后可反复筛选，无需重新上传四份素材。</div></div><button class="filter-reset" id="filterReset" type="button">清空筛选</button></div>
  <div class="filter-grid" id="filterGrid"></div>
  <div class="muted filter-note">客户编码、CP、产品名称，以及产品组的 S3 → S2 → S1 → Brand 层级路径作为隐藏关联字段保留。</div>
</div>
<div class="filter-kpis">
  <div class="filter-kpi"><span>匹配明细</span><b id="kpiRows">0</b></div>
  <div class="filter-kpi"><span>历史HARD</span><b id="kpiHard">0</b></div>
  <div class="filter-kpi"><span>本月已实现</span><b id="kpiActual">0</b></div>
  <div class="filter-kpi"><span>本月预测上限</span><b id="kpiForecast">0</b></div>
  <div class="filter-kpi"><span>未来SOFT</span><b id="kpiSoft">0</b></div>
  <div class="filter-kpi"><span>可见上限</span><b id="kpiTotal">0</b></div>
</div>
<div class="p filter-table-wrap">
  <div class="filter-summary-title"><h3>筛选汇总</h3><div class="muted" id="summaryCount"></div></div>
  <table class="filter-table"><thead><tr><th>客户</th><th class="activity">活动</th><th class="num">余额</th><th class="num">历史HARD</th><th class="num">本月已实现</th><th class="num">本月预测上限</th><th class="num">未来SOFT</th><th class="num">可见上限</th><th>状态</th></tr></thead><tbody id="summaryBody"></tbody></table>
</div>
<div class="p filter-table-wrap filter-detail">
  <div class="filter-summary-title"><h3>筛选明细</h3><div class="muted" id="detailCount"></div></div>
  <table><thead><tr><th>客户</th><th class="activity">活动</th><th>月份</th><th>渠道</th><th>Brand</th><th>产品组</th><th>SKU</th><th class="num">SO</th><th class="num">预测SO</th><th class="num">费率</th><th class="num">净承载</th><th class="num">预测承载</th><th class="num">可见承载</th></tr></thead><tbody id="detailBody"></tbody></table>
</div>
<script>
(function(){
  var ROWS=${safeRows};
  var CUR=${currentMonth};
  var defs=[
    {key:"customerKey",label:"客户",text:"customer"},
    {key:"cp",label:"活动",text:"activity"},
    {key:"month",label:"月份",text:"month"},
    {key:"channel",label:"渠道",text:"channel"},
    {key:"brand",label:"Brand",text:"brand"},
    {key:"sku",label:"SKU",text:"sku"},
    {key:"productGroup",label:"产品组",text:"productGroup"}
  ];
  var selected={}; defs.forEach(function(d){selected[d.key]=new Set();});
  var moneyFmt=new Intl.NumberFormat("zh-CN",{maximumFractionDigits:0});
  var pctFmt=new Intl.NumberFormat("zh-CN",{style:"percent",minimumFractionDigits:0,maximumFractionDigits:2});
  function money(v){return moneyFmt.format(Number(v)||0);}
  function monthLabel(v){if(!/^20\\d{4}$/.test(v))return v;return v.slice(0,4)+"年"+Number(v.slice(4))+"月";}
  function labelFor(d,r){return d.key==="month"?monthLabel(r[d.text]):String(r[d.text]||"");}
  function optionMap(d){var m=new Map();ROWS.forEach(function(r){var v=String(r[d.key]||"");if(!v)return;if(!m.has(v))m.set(v,labelFor(d,r));});return Array.from(m.entries()).sort(function(a,b){return String(a[1]).localeCompare(String(b[1]),"zh-CN",{numeric:true});});}
  function updateSummary(d,summary){var n=selected[d.key].size;summary.textContent=d.label+" · "+(n?n+"项":"全部");}
  function buildFilters(){
    var grid=document.getElementById("filterGrid");
    defs.forEach(function(d){
      var box=document.createElement("details");box.className="filter-dd";
      var summary=document.createElement("summary");box.appendChild(summary);updateSummary(d,summary);
      var menu=document.createElement("div");menu.className="filter-menu";
      optionMap(d).forEach(function(pair){
        var lab=document.createElement("label");lab.className="filter-option";
        var cb=document.createElement("input");cb.type="checkbox";cb.value=pair[0];
        var text=document.createElement("span");text.textContent=pair[1];
        cb.addEventListener("change",function(){if(cb.checked)selected[d.key].add(cb.value);else selected[d.key].delete(cb.value);updateSummary(d,summary);apply();});
        lab.appendChild(cb);lab.appendChild(text);menu.appendChild(lab);
      });
      box.appendChild(menu);grid.appendChild(box);
    });
  }
  function matches(r){return defs.every(function(d){var s=selected[d.key];return s.size===0||s.has(String(r[d.key]||""));});}
  function aggregate(rows){
    var map=new Map();
    rows.forEach(function(r){
      var key=r.customerKey+"\\u241e"+r.cp;
      var g=map.get(key);if(!g){g={customer:r.customer,customerCode:r.customerCode,activity:r.activity,cp:r.cp,balance:r.balance,months:new Map(),unknown:false};map.set(key,g);}
      var m=g.months.get(r.month);if(!m){m={actual:0,predicted:0};g.months.set(r.month,m);}
      m.actual+=Number(r.net)||0;m.predicted+=Number(r.forecastCapacity)||0;if(r.rate==null)g.unknown=true;
    });
    return Array.from(map.values()).map(function(g){
      var hard=0,actual=0,forecast=0,soft=0;
      g.months.forEach(function(v,mm){var a=v.actual>=500?v.actual:0;var p=v.predicted>=500?v.predicted:0;if(mm<CUR)hard+=a;else if(mm===CUR){actual=a;forecast=p;}else soft+=p;});
      g.hard=hard;g.actual=actual;g.forecast=forecast;g.soft=soft;g.total=Math.min(g.balance,hard+forecast+soft);return g;
    }).sort(function(a,b){return b.total-a.total||String(a.customer).localeCompare(String(b.customer),"zh-CN");});
  }
  function td(text,cls){var x=document.createElement("td");if(cls)x.className=cls;x.textContent=text;return x;}
  function renderSummary(groups){
    var body=document.getElementById("summaryBody");body.textContent="";
    groups.forEach(function(g){var tr=document.createElement("tr");tr.dataset.customerCode=g.customerCode;tr.dataset.cp=g.cp;tr.appendChild(td(g.customer));tr.appendChild(td(g.activity,"activity"));tr.appendChild(td(money(g.balance),"num"));tr.appendChild(td(money(g.hard),"num"));tr.appendChild(td(money(g.actual),"num"));tr.appendChild(td(money(g.forecast),"num"));tr.appendChild(td(money(g.soft),"num"));var total=td(money(g.total),"num");var b=document.createElement("b");b.textContent=total.textContent;total.textContent="";total.appendChild(b);tr.appendChild(total);var st=td(g.unknown?"有费率待确认":"可测算");if(g.unknown)st.className="bad";tr.appendChild(st);body.appendChild(tr);});
    if(!groups.length){var tr=document.createElement("tr");var x=td("当前筛选无结果");x.colSpan=9;x.className="filter-empty";tr.appendChild(x);body.appendChild(tr);}
    document.getElementById("summaryCount").textContent=groups.length+" 个客户×活动组合";
  }
  function renderDetail(rows){
    var body=document.getElementById("detailBody");body.textContent="";var max=300;
    rows.slice(0,max).forEach(function(r){var tr=document.createElement("tr");tr.dataset.customerCode=r.customerCode;tr.dataset.cp=r.cp;tr.dataset.productName=r.productName;tr.dataset.productPath=r.groupPath;tr.appendChild(td(r.customer));tr.appendChild(td(r.activity,"activity"));tr.appendChild(td(monthLabel(r.month)));tr.appendChild(td(r.channel));tr.appendChild(td(r.brand));tr.appendChild(td(r.productGroup));tr.appendChild(td(r.sku));tr.appendChild(td(money(r.so),"num"));tr.appendChild(td(money(r.forecastSO),"num"));tr.appendChild(td(r.rate==null?"待确认":pctFmt.format(r.rate),"num"));tr.appendChild(td(r.net==null?"—":money(r.net),"num"));tr.appendChild(td(r.forecastCapacity==null?"—":money(r.forecastCapacity),"num"));tr.appendChild(td(r.visibleCapacity==null?"—":money(r.visibleCapacity),"num"));body.appendChild(tr);});
    if(!rows.length){var tr=document.createElement("tr");var x=td("当前筛选无结果");x.colSpan=13;x.className="filter-empty";tr.appendChild(x);body.appendChild(tr);}
    document.getElementById("detailCount").textContent=rows.length>max?"共 "+rows.length+" 行，页面展示前 "+max+" 行":"共 "+rows.length+" 行";
  }
  function apply(){
    var rows=ROWS.filter(matches);var groups=aggregate(rows);renderSummary(groups);renderDetail(rows);
    var sums=groups.reduce(function(a,g){a.hard+=g.hard;a.actual+=g.actual;a.forecast+=g.forecast;a.soft+=g.soft;a.total+=g.total;return a;},{hard:0,actual:0,forecast:0,soft:0,total:0});
    document.getElementById("kpiRows").textContent=money(rows.length);document.getElementById("kpiHard").textContent=money(sums.hard);document.getElementById("kpiActual").textContent=money(sums.actual);document.getElementById("kpiForecast").textContent=money(sums.forecast);document.getElementById("kpiSoft").textContent=money(sums.soft);document.getElementById("kpiTotal").textContent=money(sums.total);
  }
  document.getElementById("filterReset").addEventListener("click",function(){defs.forEach(function(d){selected[d.key].clear();});document.querySelectorAll("#filterGrid input[type=checkbox]").forEach(function(x){x.checked=false;});document.querySelectorAll("#filterGrid details").forEach(function(box,i){updateSummary(defs[i],box.querySelector("summary"));box.open=false;});apply();});
  document.addEventListener("click",function(e){document.querySelectorAll(".filter-dd[open]").forEach(function(box){if(!box.contains(e.target))box.open=false;});});
  buildFilters();apply();
})();
</script>`;
}
