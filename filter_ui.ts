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
      rate: x["费率"] == null ? null : Number(x["费率"]),
      net: x["净承载"] == null ? null : Number(x["净承载"]),
      forecastCapacity: x["预测承载"] == null ? null : Number(x["预测承载"]),
      visibleCapacity: x["计入可见承载"] == null ? null : Number(x["计入可见承载"]),
    };
  });

  const safeRows = JSON.stringify(rows).replace(/</g, "\\u003c").replace(/-->/g, "--\\>");
  const currentMonth = JSON.stringify(String(r.currentMonth || ""));

  return `<style>
.rf{margin:14px 0;padding:16px}.rf-head{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:14px}.rf-head h3{margin:0}.rf-reset{border:1px solid #cfd8e6;background:#fff;color:#1769ff;border-radius:8px;padding:8px 12px;cursor:pointer}.rf-grid{display:grid;grid-template-columns:repeat(3,minmax(240px,1fr));gap:12px}.rf-card{border:1px solid #dbe3ef;border-radius:11px;background:#fff;padding:10px;min-height:190px}.rf-title{display:flex;justify-content:space-between;gap:8px;font-weight:700;margin-bottom:8px}.rf-title .muted{font-size:12px;font-weight:400}.rf-list{max-height:145px;overflow:auto;border-top:1px solid #edf0f5;padding-top:6px}.rf-opt{display:flex;align-items:flex-start;gap:8px;padding:6px 4px;border-radius:7px;cursor:pointer}.rf-opt:hover{background:#f4f7fb}.rf-opt span{line-height:1.35;word-break:break-word}.rf-wide{grid-column:1/-1;min-height:auto}.rf-levels{display:flex;gap:8px;flex-wrap:wrap;margin:2px 0 10px}.rf-levels label{display:inline-flex;align-items:center;gap:5px;border:1px solid #dbe3ef;border-radius:999px;padding:5px 10px;cursor:pointer}.rf-levels label.on{border-color:#1769ff;background:#eef5ff;color:#1769ff}.rf-group-list{display:grid;grid-template-columns:repeat(3,minmax(220px,1fr));gap:2px 12px;max-height:170px;overflow:auto;border-top:1px solid #edf0f5;padding-top:7px}.rf-months{display:flex;flex-wrap:wrap;gap:8px}.rf-month{display:inline-flex;align-items:center;gap:6px;border:1px solid #dbe3ef;border-radius:9px;padding:7px 10px;cursor:pointer}.rf-month.on{border-color:#1769ff;background:#eef5ff;color:#1769ff}.rf-note{margin-top:10px}.rf-kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin:14px 0}.rf-kpi{background:#fff;border:1px solid #dbe3ef;border-radius:12px;padding:12px}.rf-kpi span{display:block;color:#6b7688;font-size:12px}.rf-kpi b{display:block;font-size:21px}.rf-table{overflow:auto}.rf-table table{min-width:1220px}.rf-bar{display:flex;justify-content:space-between;align-items:center;gap:12px}.rf-bar h3{margin:0}.rf-metrics{display:flex;gap:7px;flex-wrap:wrap}.rf-metrics button{border:1px solid #cfd8e6;background:#fff;border-radius:8px;padding:6px 10px;cursor:pointer}.rf-metrics button.on{border-color:#1769ff;background:#eef5ff;color:#1769ff;font-weight:600}.rf-sub{display:block;color:#98a2b3;font-size:12px;margin-top:2px}.rf-empty{padding:24px;text-align:center;color:#6b7688}.rf-detail{margin-top:14px}.rf-detail table{min-width:1100px}.rf-detail th,.rf-detail td{vertical-align:top}.rf-activity{min-width:250px}.rf-sku{min-width:250px}.rf-month-head{min-width:88px;text-align:right}@media(max-width:1100px){.rf-grid{grid-template-columns:repeat(2,minmax(230px,1fr))}.rf-wide{grid-column:1/-1}.rf-group-list{grid-template-columns:repeat(2,minmax(220px,1fr))}.rf-kpis{grid-template-columns:repeat(3,1fr)}}@media(max-width:700px){.rf-grid{grid-template-columns:1fr}.rf-wide{grid-column:1}.rf-group-list{grid-template-columns:1fr}.rf-kpis{grid-template-columns:repeat(2,1fr)}}
</style>
<div class="p rf">
  <div class="rf-head"><div><h3>结果筛选</h3><div class="muted">客户代码是客户名称下的二级筛选；一次计算后可反复筛选。</div></div><button id="rfReset" class="rf-reset" type="button">清空筛选</button></div>
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
<div class="rf-kpis">
  <div class="rf-kpi"><span>匹配明细</span><b id="kRows">0</b></div><div class="rf-kpi"><span>历史HARD</span><b id="kHard">0</b></div><div class="rf-kpi"><span>本月已实现</span><b id="kActual">0</b></div><div class="rf-kpi"><span>本月预测上限</span><b id="kForecast">0</b></div><div class="rf-kpi"><span>未来SOFT</span><b id="kSoft">0</b></div><div class="rf-kpi"><span>可见上限</span><b id="kTotal">0</b></div>
</div>
<div class="p rf-table"><div class="rf-bar"><h3>筛选汇总</h3><div class="muted" id="summaryCount"></div></div><table><thead><tr><th>客户</th><th class="rf-activity">活动</th><th class="num">余额</th><th class="num">历史HARD</th><th class="num">本月已实现</th><th class="num">本月预测上限</th><th class="num">未来SOFT</th><th class="num">可见上限</th><th>状态</th></tr></thead><tbody id="summaryBody"></tbody></table></div>
<div class="p rf-table rf-detail"><div class="rf-bar"><div><h3>筛选明细</h3><div class="muted" id="detailCount"></div></div><div class="rf-metrics" id="metricSwitch"></div></div><table><thead id="detailHead"></thead><tbody id="detailBody"></tbody></table></div>
<script>
(function(){
  var R=${safeRows}, CUR=${currentMonth};
  var S={customer:new Set(),customerCode:new Set(),cp:new Set(),channel:new Set(),brand:new Set(),sku:new Set(),group:new Set(),month:new Set()};
  var groupLevel="s3", metric="visibleCapacity";
  var levels=[{key:"brand",label:"Brand"},{key:"s1",label:"S1"},{key:"s2",label:"S2"},{key:"s3",label:"S3"}];
  var metrics=[{key:"visibleCapacity",label:"可见承载"},{key:"net",label:"净承载"},{key:"forecastCapacity",label:"预测承载"},{key:"so",label:"SO"}];
  var fmt=new Intl.NumberFormat("zh-CN",{maximumFractionDigits:0});
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
  function metricButtons(){var b=document.getElementById("metricSwitch");b.textContent="";metrics.forEach(function(d){var x=document.createElement("button");x.type="button";x.textContent=d.label;if(metric===d.key)x.className="on";x.onclick=function(){metric=d.key;metricButtons();apply()};b.appendChild(x)})}
  function match(r){return (!S.customer.size||S.customer.has(r.customer))&&(!S.customerCode.size||S.customerCode.has(r.customerCode))&&(!S.cp.size||S.cp.has(r.cp))&&(!S.channel.size||S.channel.has(r.channel))&&(!S.brand.size||S.brand.has(r.brand))&&(!S.sku.size||S.sku.has(r.sku))&&(!S.group.size||S.group.has(String(r[groupLevel]||"")))&&(!S.month.size||S.month.has(r.month))}
  function aggregate(rows){var M=new Map();rows.forEach(function(r){var k=[r.customer,r.customerCode,r.cp].join("\\u241e"),g=M.get(k);if(!g){g={customer:r.customer,customerCode:r.customerCode,cp:r.cp,activity:r.activity,balance:r.balance,months:new Map(),unknown:false};M.set(k,g)}var m=g.months.get(r.month)||{actual:0,predicted:0};m.actual+=Number(r.net)||0;m.predicted+=Number(r.forecastCapacity)||0;g.months.set(r.month,m);if(r.rate==null)g.unknown=true});return Array.from(M.values()).map(function(g){var hard=0,actual=0,forecast=0,soft=0;g.months.forEach(function(v,mm){var a=v.actual>=500?v.actual:0,p=v.predicted>=500?v.predicted:0;if(mm<CUR)hard+=a;else if(mm===CUR){actual=a;forecast=p}else soft+=p});g.hard=hard;g.actual=actual;g.forecast=forecast;g.soft=soft;g.total=Math.min(g.balance,hard+forecast+soft);return g}).sort(function(a,b){return b.total-a.total||cmp(a.customer,b.customer)})}
  function td(t,cls){var x=document.createElement("td");if(cls)x.className=cls;x.textContent=t;return x}
  function two(main,sub,cls){var x=document.createElement("td"),b=document.createElement("b");if(cls)x.className=cls;b.textContent=main;x.appendChild(b);if(sub){var s=document.createElement("span");s.className="rf-sub";s.textContent=sub;x.appendChild(s)}return x}
  function summary(groups){var b=document.getElementById("summaryBody");b.textContent="";groups.forEach(function(g){var tr=document.createElement("tr");tr.appendChild(two(g.customer,g.customerCode));tr.appendChild(two(g.cp,g.activity,"rf-activity"));tr.appendChild(td(money(g.balance),"num"));tr.appendChild(td(money(g.hard),"num"));tr.appendChild(td(money(g.actual),"num"));tr.appendChild(td(money(g.forecast),"num"));tr.appendChild(td(money(g.soft),"num"));tr.appendChild(td(money(g.total),"num"));var st=td(g.unknown?"有费率待确认":"可测算");if(g.unknown)st.className="bad";tr.appendChild(st);b.appendChild(tr)});if(!groups.length){var tr=document.createElement("tr"),x=td("当前筛选无结果");x.colSpan=9;x.className="rf-empty";tr.appendChild(x);b.appendChild(tr)}document.getElementById("summaryCount").textContent=groups.length+" 个客户代码×活动组合"}
  function metricValue(r){var v=r[metric];return v==null?0:Number(v)||0}
  function detail(rows){var mm=uniq(rows.map(function(r){return r.month})),H=document.getElementById("detailHead"),B=document.getElementById("detailBody"),M=new Map();H.textContent="";B.textContent="";var hr=document.createElement("tr");["客户","活动","渠道","产品组","SKU"].forEach(function(v){var th=document.createElement("th");th.textContent=v;hr.appendChild(th)});mm.forEach(function(v){var th=document.createElement("th");th.className="num rf-month-head";th.textContent=mlabel(v);hr.appendChild(th)});H.appendChild(hr);rows.forEach(function(r){var group=String(r[groupLevel]||""),k=[r.customer,r.customerCode,r.cp,r.channel,group,r.sku].join("\\u241e"),g=M.get(k);if(!g){g={customer:r.customer,customerCode:r.customerCode,cp:r.cp,activity:r.activity,channel:r.channel,group:group,sku:r.sku,productName:r.productName,months:new Map()};M.set(k,g)}g.months.set(r.month,(g.months.get(r.month)||0)+metricValue(r))});var a=Array.from(M.values()).sort(function(x,y){return cmp(x.customer,y.customer)||cmp(x.cp,y.cp)||cmp(x.sku,y.sku)}),max=300;a.slice(0,max).forEach(function(g){var tr=document.createElement("tr");tr.appendChild(two(g.customer,g.customerCode));tr.appendChild(two(g.cp,g.activity,"rf-activity"));tr.appendChild(td(g.channel));tr.appendChild(td(g.group));tr.appendChild(two(g.sku,g.productName,"rf-sku"));mm.forEach(function(v){tr.appendChild(td(money(g.months.get(v)||0),"num"))});B.appendChild(tr)});if(!a.length){var tr=document.createElement("tr"),x=td("当前筛选无结果");x.colSpan=5+mm.length;x.className="rf-empty";tr.appendChild(x);B.appendChild(tr)}var ml=metrics.find(function(d){return d.key===metric}).label;document.getElementById("detailCount").textContent=(a.length>max?"共 "+a.length+" 行，展示前 "+max+" 行":"共 "+a.length+" 行")+" · 当前指标："+ml}
  function counts(){document.getElementById("nCustomer").textContent=S.customer.size?S.customer.size+"项":"全部";document.getElementById("nCP").textContent=S.cp.size?S.cp.size+"项":"全部";document.getElementById("nChannel").textContent=S.channel.size?S.channel.size+"项":"全部";document.getElementById("nBrand").textContent=S.brand.size?S.brand.size+"项":"全部";document.getElementById("nSKU").textContent=S.sku.size?S.sku.size+"项":"全部";if(S.customer.size)document.getElementById("nCode").textContent=S.customerCode.size?S.customerCode.size+"项":uniq(codeSource().map(function(r){return r.customerCode})).length+"个可选"}
  function apply(){var rows=R.filter(match),g=aggregate(rows);summary(g);detail(rows);var z=g.reduce(function(a,x){a.h+=x.hard;a.a+=x.actual;a.f+=x.forecast;a.s+=x.soft;a.t+=x.total;return a},{h:0,a:0,f:0,s:0,t:0});document.getElementById("kRows").textContent=money(rows.length);document.getElementById("kHard").textContent=money(z.h);document.getElementById("kActual").textContent=money(z.a);document.getElementById("kForecast").textContent=money(z.f);document.getElementById("kSoft").textContent=money(z.s);document.getElementById("kTotal").textContent=money(z.t);counts()}
  function reset(){Object.keys(S).forEach(function(k){S[k].clear()});groupLevel="s3";metric="visibleCapacity";customers();codes();simple("cp","fCP","nCP");simple("channel","fChannel","nChannel");simple("brand","fBrand","nBrand");simple("sku","fSKU","nSKU");groupLevels();groups();months();metricButtons();apply()}
  customers();codes();simple("cp","fCP","nCP");simple("channel","fChannel","nChannel");simple("brand","fBrand","nBrand");simple("sku","fSKU","nSKU");groupLevels();groups();months();metricButtons();apply();document.getElementById("rfReset").onclick=reset;
})();
</script>`;
}
