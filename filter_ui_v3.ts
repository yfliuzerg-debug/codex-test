export function renderFilterDashboard(r: any): string {
  const meta = new Map<string, any>();
  for (const s of r.so || []) {
    const sku = String(s.sku || "");
    if (!sku) continue;
    const old = meta.get(sku) || {};
    meta.set(sku, {
      s1: old.s1 || String(s.s1 || ""),
      s2: old.s2 || String(s.s2 || ""),
      s3: old.s3 || String(s.s3 || ""),
      brand: old.brand || String(s.brand || ""),
      productName: old.productName || String(s.skuName || ""),
    });
  }

  const rows = (r.exportRows || []).map((x: any) => {
    const sku = String(x.SKU || "");
    const m = meta.get(sku) || {};
    const customer = String(x["经销商/客户"] || "");
    const customerCode = String(x["客户编码"] || "");
    const cp = String(x.CP || "");
    const activity = String(x["活动"] || "");
    const productName = String(x["产品名称"] || m.productName || "");
    const brand = String(x["品牌"] || m.brand || "");
    return {
      customer,
      customerCode,
      cp,
      activity,
      activityLabel: [cp, activity].filter(Boolean).join("｜"),
      month: String(x["月份"] || ""),
      channel: String(x["渠道"] || ""),
      brand,
      s1: String(m.s1 || ""),
      s2: String(m.s2 || ""),
      s3: String(m.s3 || ""),
      sku,
      productName,
      skuLabel: [sku, productName].filter(Boolean).join("｜"),
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
.rf{margin:14px 0;padding:16px}.rf-head{display:flex;justify-content:space-between;align-items:center;gap:12px}.rf-head h3{margin:0}.rf-actions{display:flex;gap:8px;flex-wrap:wrap}.rf-btn{border:1px solid #cfd8e6;background:#fff;color:#1769ff;border-radius:8px;padding:8px 12px;cursor:pointer}.rf-btn.primary{background:#1769ff;color:#fff;border-color:#1769ff}.rf-btn:disabled{opacity:.55;cursor:not-allowed}.rf-body{margin-top:14px}.rf.collapsed .rf-body{display:none}.rf-state{margin-top:2px}.rf-grid{display:grid;grid-template-columns:repeat(3,minmax(240px,1fr));gap:12px}.rf-card{border:1px solid #dbe3ef;border-radius:11px;background:#fff;padding:10px;min-height:190px}.rf-title{display:flex;justify-content:space-between;gap:8px;font-weight:700;margin-bottom:8px}.rf-title .muted{font-size:12px;font-weight:400}.rf-list{max-height:145px;overflow:auto;border-top:1px solid #edf0f5;padding-top:6px}.rf-opt{display:flex;align-items:flex-start;gap:8px;padding:6px 4px;border-radius:7px;cursor:pointer}.rf-opt:hover{background:#f4f7fb}.rf-opt span{line-height:1.35;word-break:break-word}.rf-wide{grid-column:1/-1;min-height:auto}.rf-levels{display:flex;gap:8px;flex-wrap:wrap;margin:2px 0 10px}.rf-levels label{display:inline-flex;align-items:center;gap:5px;border:1px solid #dbe3ef;border-radius:999px;padding:5px 10px;cursor:pointer}.rf-levels label.on{border-color:#1769ff;background:#eef5ff;color:#1769ff}.rf-group-list{display:grid;grid-template-columns:repeat(3,minmax(220px,1fr));gap:2px 12px;max-height:170px;overflow:auto;border-top:1px solid #edf0f5;padding-top:7px}.rf-months{display:flex;flex-wrap:wrap;gap:8px}.rf-month{display:inline-flex;align-items:center;gap:6px;border:1px solid #dbe3ef;border-radius:9px;padding:7px 10px;cursor:pointer}.rf-month.on{border-color:#1769ff;background:#eef5ff;color:#1769ff}.rf-note{margin-top:10px}.rf-kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin:14px 0}.rf-kpi{background:#fff;border:1px solid #dbe3ef;border-radius:12px;padding:12px}.rf-kpi span{display:block;color:#6b7688;font-size:12px}.rf-kpi b{display:block;font-size:21px}.rf-table{overflow:auto}.rf-table table{min-width:1220px}.rf-bar{display:flex;justify-content:space-between;align-items:center;gap:12px}.rf-bar h3{margin:0}.rf-metrics{display:flex;gap:7px;flex-wrap:wrap}.rf-metrics button{border:1px solid #cfd8e6;background:#fff;border-radius:8px;padding:6px 10px;cursor:pointer}.rf-metrics button.on{border-color:#1769ff;background:#eef5ff;color:#1769ff;font-weight:600}.rf-sub{display:block;color:#98a2b3;font-size:12px;margin-top:2px}.rf-empty{padding:24px;text-align:center;color:#6b7688}.rf-detail{margin-top:14px}.rf-detail table{min-width:1080px}.rf-detail th,.rf-detail td{vertical-align:top}.rf-activity{min-width:260px}.rf-sku{min-width:280px}.rf-month-head{min-width:88px;text-align:right}@media(max-width:1100px){.rf-grid{grid-template-columns:repeat(2,minmax(230px,1fr))}.rf-wide{grid-column:1/-1}.rf-group-list{grid-template-columns:repeat(2,minmax(220px,1fr))}.rf-kpis{grid-template-columns:repeat(3,1fr)}}@media(max-width:700px){.rf-grid{grid-template-columns:1fr}.rf-wide{grid-column:1}.rf-group-list{grid-template-columns:1fr}.rf-kpis{grid-template-columns:repeat(2,1fr)}}
</style>
<div class="p rf" id="rfPanel">
  <div class="rf-head">
    <div><h3>结果筛选</h3><div class="muted rf-state" id="rfState">客户代码是客户名称下的二级筛选；一次计算后可反复筛选。</div></div>
    <div class="rf-actions"><button id="rfExport" class="rf-btn primary" type="button">导出筛选结果.xlsx</button><button id="rfReset" class="rf-btn" type="button">清空筛选</button><button id="rfToggle" class="rf-btn" type="button">收起筛选</button></div>
  </div>
  <div class="rf-body" id="rfBody">
    <div class="rf-grid">
      <div class="rf-card"><div class="rf-title"><span>客户</span><span class="muted" id="nCustomer"></span></div><div class="rf-list" id="fCustomer"></div></div>
      <div class="rf-card"><div class="rf-title"><span>客户代码</span><span class="muted" id="nCode"></span></div><div class="rf-list" id="fCode"></div></div>
      <div class="rf-card"><div class="rf-title"><span>活动</span><span class="muted" id="nCP"></span></div><div class="rf-list" id="fCP"></div></div>
      <div class="rf-card"><div class="rf-title"><span>渠道</span><span class="muted" id="nChannel"></span></div><div class="rf-list" id="fChannel"></div></div>
      <div class="rf-card"><div class="rf-title"><span>Brand</span><span class="muted" id="nBrand"></span></div><div class="rf-list" id="fBrand"></div></div>
      <div class="rf-card"><div class="rf-title"><span>SKU</span><span class="muted" id="nSKU"></span></div><div class="rf-list" id="fSKU"></div></div>
      <div class="rf-card rf-wide"><div class="rf-title"><span>产品组</span><span class="muted" id="nGroup"></span></div><div class="rf-levels" id="groupLevels"></div><div class="rf-group-list" id="fGroup"></div></div>
      <div class="rf-card rf-wide"><div class="rf-title"><span>月份</span><span class="muted" id="nMonth"></span></div><div class="rf-months" id="fMonth"></div></div>
    </div>
    <div class="muted rf-note">活动显示 CP｜活动名称；SKU 显示 SKU编码｜产品名称；产品组按 Brand → S1 → S2 → S3 单层切换。</div>
  </div>
</div>
<div class="rf-kpis">
  <div class="rf-kpi"><span>匹配明细</span><b id="kRows">0</b></div><div class="rf-kpi"><span>历史HARD</span><b id="kHard">0</b></div><div class="rf-kpi"><span>本月已实现</span><b id="kActual">0</b></div><div class="rf-kpi"><span>本月预测上限</span><b id="kForecast">0</b></div><div class="rf-kpi"><span>未来SOFT</span><b id="kSoft">0</b></div><div class="rf-kpi"><span>可见上限</span><b id="kTotal">0</b></div>
</div>
<div class="p rf-table"><div class="rf-bar"><h3>筛选汇总</h3><div class="muted" id="summaryCount"></div></div><table><thead><tr><th>客户</th><th class="rf-activity">活动</th><th class="num">余额</th><th class="num">历史HARD</th><th class="num">本月已实现</th><th class="num">本月预测上限</th><th class="num">未来SOFT</th><th class="num">可见上限</th><th>状态</th></tr></thead><tbody id="summaryBody"></tbody></table></div>
<div class="p rf-table rf-detail"><div class="rf-bar"><div><h3>筛选明细</h3><div class="muted" id="detailCount"></div></div><div class="rf-metrics" id="metricSwitch"></div></div><table><thead id="detailHead"></thead><tbody id="detailBody"></tbody></table></div>
<script>
(function(){
  var R=${safeRows}, CUR=${currentMonth};
  var S={customer:new Set(),customerCode:new Set(),cp:new Set(),channel:new Set(),brand:new Set(),sku:new Set(),group:new Set(),month:new Set()};
  var groupLevel="s3", metric="visibleCapacity", collapsed=false;
  var levels=[{key:"brand",label:"Brand"},{key:"s1",label:"S1"},{key:"s2",label:"S2"},{key:"s3",label:"S3"}];
  var metrics=[{key:"visibleCapacity",label:"可见承载"},{key:"net",label:"净承载"},{key:"forecastCapacity",label:"预测承载"},{key:"so",label:"SO"}];
  var fmt=new Intl.NumberFormat("zh-CN",{maximumFractionDigits:0});
  var lastRows=[],lastGroups=[],lastPivot=[],lastMonths=[];
  function money(v){return fmt.format(Number(v)||0)}
  function mlabel(v){return /^20\\d{4}$/.test(v)?Number(v.slice(4))+"月":v}
  function cmp(a,b){return String(a).localeCompare(String(b),"zh-CN",{numeric:true})}
  function uniq(a){return Array.from(new Set(a.filter(function(v){return String(v||"")!==""}))).sort(cmp)}
  function opt(text,value,set,change){var l=document.createElement("label"),c=document.createElement("input"),s=document.createElement("span");l.className="rf-opt";c.type="checkbox";c.value=value;c.checked=set.has(value);s.textContent=text;c.onchange=function(){c.checked?set.add(value):set.delete(value);change()};l.appendChild(c);l.appendChild(s);return l}
  function pairs(key){var m=new Map();R.forEach(function(r){var v=String(r[key]||"");if(!v||m.has(v))return;m.set(v,key==="cp"?r.activityLabel:key==="sku"?r.skuLabel:v)});return Array.from(m.entries()).sort(function(a,b){return cmp(a[1],b[1])})}
  function simple(key,boxId,numId){var b=document.getElementById(boxId);b.textContent="";pairs(key).forEach(function(p){b.appendChild(opt(p[1],p[0],S[key],apply))});document.getElementById(numId).textContent=S[key].size?S[key].size+"项":"全部"}
  function customers(){var b=document.getElementById("fCustomer");b.textContent="";uniq(R.map(function(r){return r.customer})).forEach(function(v){b.appendChild(opt(v,v,S.customer,function(){syncCodes();apply()}))});document.getElementById("nCustomer").textContent=S.customer.size?S.customer.size+"项":"全部"}
  function codeSource(){return S.customer.size?R.filter(function(r){return S.customer.has(r.customer)}):[]}
  function syncCodes(){var allowed=new Set(codeSource().map(function(r){return r.customerCode}).filter(Boolean));Array.from(S.customerCode).forEach(function(v){if(!allowed.has(v))S.customerCode.delete(v)});codes()}
  function codes(){var b=document.getElementById("fCode");b.textContent="";if(!S.customer.size){var x=document.createElement("div");x.className="muted";x.style.padding="8px 4px";x.textContent="先选择客户";b.appendChild(x);document.getElementById("nCode").textContent="二级筛选";return}var a=uniq(codeSource().map(function(r){return r.customerCode}));a.forEach(function(v){b.appendChild(opt(v,v,S.customerCode,apply))});document.getElementById("nCode").textContent=S.customerCode.size?S.customerCode.size+"项":a.length+"个可选"}
  function groupLevels(){var b=document.getElementById("groupLevels");b.textContent="";levels.forEach(function(d){var l=document.createElement("label"),c=document.createElement("input"),s=document.createElement("span");if(groupLevel===d.key)l.className="on";c.type="radio";c.name="groupLevel";c.checked=groupLevel===d.key;s.textContent=d.label;c.onchange=function(){if(!c.checked)return;groupLevel=d.key;S.group.clear();groupLevels();groups();apply()};l.appendChild(c);l.appendChild(s);b.appendChild(l)})}
  function groups(){var b=document.getElementById("fGroup");b.textContent="";uniq(R.map(function(r){return r[groupLevel]})).forEach(function(v){b.appendChild(opt(v,v,S.group,apply))});var lab=levels.find(function(d){return d.key===groupLevel}).label;document.getElementById("nGroup").textContent=lab+" · "+(S.group.size?S.group.size+"项":"全部")}
  function months(){var b=document.getElementById("fMonth");b.textContent="";uniq(R.map(function(r){return r.month})).forEach(function(v){var l=document.createElement("label"),c=document.createElement("input"),s=document.createElement("span");l.className="rf-month"+(S.month.has(v)?" on":"");c.type="checkbox";c.checked=S.month.has(v);s.textContent=mlabel(v);c.onchange=function(){c.checked?S.month.add(v):S.month.delete(v);months();apply()};l.appendChild(c);l.appendChild(s);b.appendChild(l)});document.getElementById("nMonth").textContent=S.month.size?S.month.size+"项":"全部"}
  function metricButtons(){var b=document.getElementById("metricSwitch");b.textContent="";metrics.forEach(function(d){var x=document.createElement("button");x.type="button";x.textContent=d.label;if(metric===d.key)x.className="on";x.onclick=function(){metric=d.key;metricButtons();renderDetail(lastRows)};b.appendChild(x)})}
  function match(r){if(S.customer.size&&!S.customer.has(r.customer))return false;if(S.customerCode.size&&!S.customerCode.has(r.customerCode))return false;if(S.cp.size&&!S.cp.has(r.cp))return false;if(S.channel.size&&!S.channel.has(r.channel))return false;if(S.brand.size&&!S.brand.has(r.brand))return false;if(S.sku.size&&!S.sku.has(r.sku))return false;if(S.group.size&&!S.group.has(r[groupLevel]))return false;if(S.month.size&&!S.month.has(r.month))return false;return true}
  function aggregate(rows){var m=new Map();rows.forEach(function(r){var k=r.customer+"\\u241e"+r.cp,g=m.get(k);if(!g){g={customer:r.customer,activity:r.activityLabel,cp:r.cp,balance:r.balance,months:new Map(),unknown:false};m.set(k,g)}var x=g.months.get(r.month);if(!x){x={actual:0,forecast:0};g.months.set(r.month,x)}x.actual+=Number(r.net)||0;x.forecast+=Number(r.forecastCapacity)||0;if(r.rate==null)g.unknown=true});return Array.from(m.values()).map(function(g){var hard=0,actual=0,forecast=0,soft=0;g.months.forEach(function(v,mm){var a=v.actual>=500?v.actual:0,p=v.forecast>=500?v.forecast:0;if(mm<CUR)hard+=a;else if(mm===CUR){actual=a;forecast=p}else soft+=p});g.hard=hard;g.actual=actual;g.forecast=forecast;g.soft=soft;g.total=Math.min(g.balance,hard+forecast+soft);return g}).sort(function(a,b){return b.total-a.total||cmp(a.customer,b.customer)})}
  function td(text,cls){var x=document.createElement("td");if(cls)x.className=cls;x.textContent=text;return x}
  function customerTd(name,code){var x=document.createElement("td"),a=document.createElement("div"),b=document.createElement("span");a.textContent=name;b.className="rf-sub";b.textContent=code;x.appendChild(a);x.appendChild(b);return x}
  function renderSummary(groups){var b=document.getElementById("summaryBody");b.textContent="";groups.forEach(function(g){var tr=document.createElement("tr");tr.appendChild(td(g.customer));tr.appendChild(td(g.activity,"rf-activity"));tr.appendChild(td(money(g.balance),"num"));tr.appendChild(td(money(g.hard),"num"));tr.appendChild(td(money(g.actual),"num"));tr.appendChild(td(money(g.forecast),"num"));tr.appendChild(td(money(g.soft),"num"));var z=td(money(g.total),"num");z.style.fontWeight="700";tr.appendChild(z);var st=td(g.unknown?"有费率待确认":"可测算");if(g.unknown)st.className="bad";tr.appendChild(st);b.appendChild(tr)});if(!groups.length){var tr=document.createElement("tr"),x=td("当前筛选无结果");x.colSpan=9;x.className="rf-empty";tr.appendChild(x);b.appendChild(tr)}document.getElementById("summaryCount").textContent=groups.length+" 个客户×活动组合"}
  function pivot(rows){var ms=S.month.size?Array.from(S.month).sort(cmp):uniq(rows.map(function(r){return r.month}));var m=new Map();rows.forEach(function(r){var gv=String(r[groupLevel]||""),k=[r.customer,r.customerCode,r.cp,r.channel,gv,r.sku].join("\\u241e"),g=m.get(k);if(!g){g={customer:r.customer,customerCode:r.customerCode,activity:r.activityLabel,channel:r.channel,group:gv,sku:r.skuLabel,values:{},unknown:false};m.set(k,g)}if(r[metric]==null)g.unknown=true;g.values[r.month]=(g.values[r.month]||0)+(Number(r[metric])||0)});var out=Array.from(m.values());if(metric!=="so")out=out.filter(function(g){return g.unknown||ms.some(function(mm){return Number(g.values[mm]||0)!==0})});out.sort(function(a,b){return cmp(a.customer,b.customer)||cmp(a.channel,b.channel)||cmp(a.group,b.group)||cmp(a.sku,b.sku)});return {months:ms,rows:out}}
  function renderDetail(rows){var p=pivot(rows),h=document.getElementById("detailHead"),b=document.getElementById("detailBody");h.textContent="";b.textContent="";var trh=document.createElement("tr");["客户","活动","渠道","产品组","SKU"].forEach(function(v){var th=document.createElement("th");th.textContent=v;if(v==="活动")th.className="rf-activity";if(v==="SKU")th.className="rf-sku";trh.appendChild(th)});p.months.forEach(function(mm){var th=document.createElement("th");th.className="rf-month-head";th.textContent=mlabel(mm);trh.appendChild(th)});h.appendChild(trh);p.rows.forEach(function(g){var tr=document.createElement("tr");tr.appendChild(customerTd(g.customer,g.customerCode));tr.appendChild(td(g.activity,"rf-activity"));tr.appendChild(td(g.channel));tr.appendChild(td(g.group));tr.appendChild(td(g.sku,"rf-sku"));p.months.forEach(function(mm){tr.appendChild(td(money(g.values[mm]||0),"num"))});b.appendChild(tr)});if(!p.rows.length){var tr=document.createElement("tr"),x=td("当前筛选无结果");x.colSpan=5+p.months.length;x.className="rf-empty";tr.appendChild(x);b.appendChild(tr)}var ml=metrics.find(function(d){return d.key===metric}).label;document.getElementById("detailCount").textContent=p.rows.length+" 行 · 当前指标："+ml;lastPivot=p.rows;lastMonths=p.months}
  function updateState(){var n=S.customer.size+S.customerCode.size+S.cp.size+S.channel.size+S.brand.size+S.sku.size+S.group.size+S.month.size;var gl=levels.find(function(d){return d.key===groupLevel}).label;document.getElementById("rfState").textContent=(collapsed?"筛选已收起 · ":"")+(n?"已选 "+n+" 项条件":"当前未限定筛选")+" · 产品组层级 "+gl}
  function apply(){lastRows=R.filter(match);lastGroups=aggregate(lastRows);renderSummary(lastGroups);renderDetail(lastRows);var sums=lastGroups.reduce(function(a,g){a.hard+=g.hard;a.actual+=g.actual;a.forecast+=g.forecast;a.soft+=g.soft;a.total+=g.total;return a},{hard:0,actual:0,forecast:0,soft:0,total:0});document.getElementById("kRows").textContent=money(lastRows.length);document.getElementById("kHard").textContent=money(sums.hard);document.getElementById("kActual").textContent=money(sums.actual);document.getElementById("kForecast").textContent=money(sums.forecast);document.getElementById("kSoft").textContent=money(sums.soft);document.getElementById("kTotal").textContent=money(sums.total);simple("cp","fCP","nCP");simple("channel","fChannel","nChannel");simple("brand","fBrand","nBrand");simple("sku","fSKU","nSKU");groups();months();customers();codes();updateState()}
  function exportSummary(){return lastGroups.map(function(g){return {"客户":g.customer,"活动":g.activity,"余额":g.balance,"历史HARD":g.hard,"本月已实现":g.actual,"本月预测上限":g.forecast,"未来SOFT":g.soft,"可见上限":g.total,"状态":g.unknown?"有费率待确认":"可测算"}})}
  function exportDetail(){return lastPivot.map(function(g){var o={"客户":g.customer,"客户编码":g.customerCode,"活动":g.activity,"渠道":g.channel,"产品组层级":levels.find(function(d){return d.key===groupLevel}).label,"产品组":g.group,"SKU":g.sku};lastMonths.forEach(function(mm){o[mlabel(mm)]=g.values[mm]||0});return o})}
  async function exportXlsx(){if(!lastRows.length){alert("当前筛选无结果，无法导出。");return}var btn=document.getElementById("rfExport"),old=btn.textContent;btn.disabled=true;btn.textContent="正在导出…";try{var ml=metrics.find(function(d){return d.key===metric}).label,res=await fetch("/export-filtered",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({summaryRows:exportSummary(),detailRows:exportDetail(),metricLabel:ml,groupLevel:groupLevel})});if(!res.ok)throw new Error(await res.text());var blob=await res.blob(),u=URL.createObjectURL(blob),a=document.createElement("a"),d=new Date(),ds=d.getFullYear()+String(d.getMonth()+1).padStart(2,"0")+String(d.getDate()).padStart(2,"0");a.href=u;a.download="费用规划助手_筛选结果_"+ds+".xlsx";document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(u)},1000)}catch(e){alert("导出失败："+(e&&e.message?e.message:String(e)))}finally{btn.disabled=false;btn.textContent=old}}
  document.getElementById("rfReset").onclick=function(){Object.keys(S).forEach(function(k){S[k].clear()});codes();apply()};
  document.getElementById("rfToggle").onclick=function(){collapsed=!collapsed;document.getElementById("rfPanel").classList.toggle("collapsed",collapsed);this.textContent=collapsed?"展开筛选":"收起筛选";this.setAttribute("aria-expanded",String(!collapsed));updateState()};
  document.getElementById("rfExport").onclick=exportXlsx;
  groupLevels();metricButtons();apply();
})();
</script>`;
}
