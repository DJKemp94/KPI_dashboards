// Shared behaviour for the university and faculty/service KPI dashboards.
// Inlined into each generated page by scripts/faculty_dashboard.py; each page
// defines facultyData, tooltipData, kpiMeta and DASHBOARD_CONFIG beforehand.
let currentSortOrder='highLow';let currentViewMode='detailed';
const activeBandFilters=new Set();
const isUniversityDashboard=DASHBOARD_CONFIG.isUniversity===true;
const INFO_ICON='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"></circle><path d="M12 16v-4"></path><path d="M12 8h.01"></path></svg>';
const BAND_INFO={
  'performance-excellent':{cls:'band-excellent',name:'Excellent',label:'95% or more'},
  'performance-good':{cls:'band-good',name:'Good',label:'80% to 94.9%'},
  'performance-warning':{cls:'band-warning',name:'Needs attention',label:'60% to 79.9%'},
  'performance-poor':{cls:'band-poor',name:'Concern',label:'Below 60%'},
  'performance-not-applicable':{cls:'band-na',name:'Not applicable',label:'Not applicable'},
  'performance-no-data':{cls:'band-na',name:'No data',label:'No data'}
};
const BAND_COLORS={excellent:'#2563eb',good:'#059669',warning:'#e08a00',poor:'#dc2626',neutral:'#94a3b8'};
function escapeHtml(value){
  return String(value===null||value===undefined?'':value)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function textToParagraphs(text){
  return String(text||'').trim().split(/\n\s*\n/).map((para)=>`<p>${escapeHtml(para.trim()).replace(/\n/g,'<br>')}</p>`).join('');
}
function formatReportDate(label){
  const match=/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(label||'').trim());
  if(!match) return String(label||'');
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const month=months[Number(match[2])-1];
  return month?`${Number(match[1])} ${month} ${match[3]}`:String(label);
}
// Small inline info button (e.g. beside "N/A"); opens the shared info dialog.
function createTooltip(kpiName,tooltipText){
  if(!tooltipText) return '';
  return `<button type="button" class="info-btn icon-only" data-info-kpi="${encodeURIComponent(kpiName)}" data-info-text="${encodeURIComponent(tooltipText)}" aria-label="More information: ${escapeHtml(kpiName)}" title="More information">${INFO_ICON}</button>`;
}
// Labelled "About" button beside each KPI title; opens the full KPI definition.
function createKpiInfoButton(kpiName,iconOnly=false){
  const label = iconOnly ? '' : '<span>About</span>';
  return `<button type="button" class="info-btn${iconOnly ? ' icon-only' : ''}" data-info-kpi="${encodeURIComponent(kpiName)}" data-info-full="1" aria-label="About this KPI: ${escapeHtml(kpiName)}" title="About this KPI">${INFO_ICON}${label}</button>`;
}
function buildKpiInfoHtml(kpiName){
  const meta=kpiMeta[kpiName]||{};
  const guidance=tooltipData[kpiName];
  let html='';
  if(meta.question) html+=`<div class="info-section"><h3>What it measures</h3><p class="info-lead">${escapeHtml(meta.question)}</p></div>`;
  if(guidance) html+=`<div class="info-section"><h3>Guidance</h3>${textToParagraphs(guidance)}</div>`;
  const facts=[];
  if(meta.denominator_label) facts.push(['Counted in',escapeHtml(meta.denominator_label.charAt(0).toUpperCase()+meta.denominator_label.slice(1))]);
  if(isEstimateKpi(kpiName)) facts.push(['Basis','Estimate reported by the H&amp;S Coordinator (not count-based)']);
  facts.push(['Not applicable when',escapeHtml(getNotApplicableText(kpiName))]);
  if(meta.zero_score_message) facts.push(['A score of 0% shows as',escapeHtml(meta.zero_score_message)]);
  html+=`<div class="info-section"><h3>Key facts</h3><dl class="info-facts">${facts.map(([k,v])=>`<dt>${k}</dt><dd>${v}</dd>`).join('')}</dl></div>`;
  return html;
}
function openInfoDialog(eyebrow,title,bodyHtml){
  const dialog=document.getElementById('infoDialog');
  document.getElementById('infoDialogEyebrow').textContent=eyebrow;
  document.getElementById('infoDialogTitle').textContent=title;
  document.getElementById('infoDialogBody').innerHTML=bodyHtml;
  if(typeof dialog.showModal==='function'){ if(!dialog.open) dialog.showModal(); }
  else dialog.setAttribute('open','');
  document.getElementById('infoDialogBody').scrollTop=0;
}
function closeInfoDialog(){
  const dialog=document.getElementById('infoDialog');
  if(typeof dialog.close==='function') dialog.close(); else dialog.removeAttribute('open');
}
function openGuide(){
  openInfoDialog('Help','How to read this dashboard',document.getElementById('guideTemplate').innerHTML);
}
let infoHandlersBound=false;
function positionTooltips(){
  if(infoHandlersBound) return;
  infoHandlersBound=true;
  document.addEventListener('click',(event)=>{
    const trigger=event.target.closest('.info-btn');
    if(!trigger) return;
    event.preventDefault();
    const kpiName=decodeURIComponent(trigger.getAttribute('data-info-kpi')||'');
    if(trigger.getAttribute('data-info-full')==='1'){
      openInfoDialog('KPI definition',kpiName,buildKpiInfoHtml(kpiName));
    } else {
      const text=decodeURIComponent(trigger.getAttribute('data-info-text')||'');
      openInfoDialog('Note',kpiName,textToParagraphs(text));
    }
  },true);
  const dialog=document.getElementById('infoDialog');
  dialog.addEventListener('click',(event)=>{ if(event.target===dialog) closeInfoDialog(); });
}
function applyCompressedValueTitles(){
  if(currentViewMode!=='compressed') return;
  document.querySelectorAll('.compressed-view .kpi-hero-wrap').forEach((valueEl)=>{
    const fullText=valueEl.textContent.replace(/\s+/g,' ').trim();
    if(fullText) valueEl.setAttribute('title',fullText);
  });
}
function isOutOfRangePercentage(percentage){if(percentage===null||percentage===undefined)return false;const numericPercentage=parseFloat(percentage);if(Number.isNaN(numericPercentage))return false;return numericPercentage<0||numericPercentage>100;}
function createOutOfRangeBadge(percentage,extraClass=''){if(!isOutOfRangePercentage(percentage))return '';const cls=extraClass?`data-warning ${extraClass}`:'data-warning';return `<span class="${cls}" title="Source percentage is outside 0-100. This can happen when plans changed or where source data may have been entered incorrectly.">!</span>`;}
function normalizeTrendHistory(history){
  if(!Array.isArray(history)) return [];
  return history
    .filter((point) => point && point.date)
    .map((point) => ({
      date: point.date,
      percentage: typeof point.percentage === 'number' ? point.percentage : null,
      rawPercentage: typeof point.raw_percentage === 'number'
        ? point.raw_percentage
        : (typeof point.rawPercentage === 'number' ? point.rawPercentage : (typeof point.percentage === 'number' ? point.percentage : null)),
      number: typeof point.number === 'number' ? point.number : null,
      applicable: point.applicable !== false,
      exceeded: point.exceeded === true,
      completed_count: typeof point.completed_count === 'number' ? point.completed_count : null
    }));
}
const exceededInferenceKpis = new Set([
  'Inspections Carried Out',
  'Leadership Walkarounds',
  'PEEPS Drilled',
  'Fire Drills Completed'
]);
function isFiniteNumber(value){
  return typeof value === 'number' && Number.isFinite(value);
}
function applyExceededInference(){
  if(!facultyData || !facultyData.kpis || !facultyData.schools) return;
  const schoolEntries = Object.entries(facultyData.schools);
  if(schoolEntries.length === 0) return;

  for(const kpiName of exceededInferenceKpis){
    const facultyKpi = facultyData.kpis[kpiName];
    if(!facultyKpi || !isFiniteNumber(facultyKpi.number) || !isFiniteNumber(facultyKpi.percentage) || facultyKpi.number <= 0) continue;
    const facultyCompleted = Math.round((facultyKpi.percentage * facultyKpi.number) / 100);

    let knownCompleted = 0;
    const zeroScopeCandidates = [];

    schoolEntries.forEach(([schoolName, schoolData]) => {
      const schoolKpi = schoolData.kpis && schoolData.kpis[kpiName];
      if(!schoolKpi) return;
      if(isMissingCurrentForKpi(kpiName, schoolKpi.percentage, schoolKpi.number, schoolKpi.applicable)) return;
      const num = Number(schoolKpi.number);
      const perc = Number(schoolKpi.percentage);
      if(!Number.isFinite(num) || !Number.isFinite(perc)) return;

      if(num > 0){
        knownCompleted += Math.round((perc * num) / 100);
        return;
      }
      if(Math.abs(num) < 1e-9 && Math.abs(perc) < 1e-9 && schoolKpi.applicable === false){
        zeroScopeCandidates.push([schoolName, schoolKpi]);
      }
    });

    const extraCompleted = facultyCompleted - knownCompleted;
    if(extraCompleted > 0 && zeroScopeCandidates.length === 1){
      const [, inferredSchoolKpi] = zeroScopeCandidates[0];
      inferredSchoolKpi.exceeded = true;
      inferredSchoolKpi.applicable = true;
      inferredSchoolKpi.completed_count = extraCompleted;
      if(!isFiniteNumber(inferredSchoolKpi.percentage)) inferredSchoolKpi.percentage = 0;

      const history = normalizeTrendHistory(inferredSchoolKpi.history);
      if(history.length > 0){
        const current = history[history.length - 1];
        current.exceeded = true;
        current.applicable = true;
        current.completed_count = extraCompleted;
        if(!isFiniteNumber(current.percentage)){
          current.percentage = isFiniteNumber(current.rawPercentage) ? current.rawPercentage : 0;
        }
        inferredSchoolKpi.history = history;
      }
    }
  }
}
function kpiUsesDenominator(kpiName){
  return !!(kpiMeta[kpiName] && kpiMeta[kpiName].denominator_label);
}
function isImplicitNoReturnPoint(kpiName, point){
  if(!point || point.applicable === false) return false;
  if(!kpiUsesDenominator(kpiName)) return false;
  const missingCount = point.number === null || point.number === undefined || Number.isNaN(point.number);
  const raw = typeof point.rawPercentage === 'number' ? point.rawPercentage : point.percentage;
  return missingCount && typeof raw === 'number' && Math.abs(raw) < 1e-9;
}
function buildTrendBounds(seriesList){
  const values = [];
  seriesList.forEach((series) => {
    if(Array.isArray(series)) {
      series.forEach((value) => {
        if(typeof value === 'number' && !Number.isNaN(value)) values.push(value);
      });
    }
  });
  if(values.length === 0) return { min: 0, max: 100 };
  let min = Math.min(...values);
  let max = Math.max(...values);
  const span = max - min;
  const padding = Math.max(4, span * 0.18);
  if(span < 8) {
    min -= 4;
    max += 4;
  } else {
    min -= padding;
    max += padding;
  }
  return {
    min: Math.floor(min),
    max: Math.ceil(max)
  };
}
function getNotApplicableText(kpiName) {
  const meta = kpiMeta[kpiName] || {};
  const custom = meta.na_message;
  if(custom) return custom;
  const label = meta.denominator_label || 'items';
  return `Not applicable (0 ${label} in scope)`;
}
function getZeroScoreText(kpiName) {
  const meta = kpiMeta[kpiName] || {};
  return meta.zero_score_message || null;
}
function isEstimateKpi(kpiName){
  return kpiName === 'Risk Assessment Coverage' || kpiName === 'Training Matrix Coverage';
}
function getInScopeTextCompressed(kpiName, countValue, schoolKpi){
  if(isEstimateKpi(kpiName)){
    return '<span class="cell-exceeded">Estimate (not count-based)</span>';
  }
  const schoolLabel = (kpiMeta[kpiName] && kpiMeta[kpiName].denominator_label) ? kpiMeta[kpiName].denominator_label : 'records';
  const hasNoScope = countValue === null;
  const isExceeded = schoolKpi.exceeded === true;
  if(isExceeded && countValue !== null && countValue === 0){
    return schoolKpi.completed_count !== null
      ? `<span class="cell-exceeded">0 scheduled, ${schoolKpi.completed_count} carried out</span>`
      : `<span class="cell-exceeded">0 scheduled</span>`;
  }
  if(hasNoScope){
    return '<span class="cell-no-return">No return</span>';
  }
  return `${Math.round(countValue)} ${schoolLabel}`;
}
function getLastApplicablePoint(cleaned){
  const points = cleaned.filter((point) => point.applicable && typeof point.percentage === 'number');
  if(points.length === 0) return null;
  return points[points.length - 1];
}
function getChangeSummary(history, kpiName){
  const cleaned = normalizeTrendHistory(history);
  if(cleaned.length === 0) return null;
  const currentPoint = cleaned[cleaned.length - 1];
  const explicitMissing = currentPoint.applicable !== false
    && (currentPoint.percentage === null || currentPoint.percentage === undefined)
    && (currentPoint.rawPercentage === null || currentPoint.rawPercentage === undefined);
  if(explicitMissing) {
    return {
      label: 'No return submitted this period',
      shortLabel: 'No return',
      className: 'change-missing',
      symbol: '!',
      delta: null
    };
  }
  if(isImplicitNoReturnPoint(kpiName, currentPoint)) {
    return {
      label: 'No return submitted this period',
      shortLabel: 'No return',
      className: 'change-missing',
      symbol: '!',
      delta: null
    };
  }
  if(currentPoint.applicable === false) {
    const priorPoint = getLastApplicablePoint(cleaned.slice(0, -1));
    const naText = getNotApplicableText(kpiName);
    if(priorPoint) {
      return {
        label: `${naText}. Last applicable return: ${priorPoint.percentage.toFixed(1)}% (${priorPoint.date}).`,
        shortLabel: `N/A | Last applicable ${priorPoint.percentage.toFixed(1)}%`,
        className: 'change-na',
        symbol: '○',
        delta: null
      };
    }
    return {
      label: `${naText}. No previous applicable return to compare.`,
      shortLabel: 'N/A | No previous applicable return',
      className: 'change-na',
      symbol: '○',
      delta: null
    };
  }
  const applicable = cleaned.filter((point) => point.applicable && typeof point.percentage === 'number' && !isImplicitNoReturnPoint(kpiName, point));
  if(applicable.length < 2) return { label: 'No previous applicable return', shortLabel: 'No previous applicable', className: 'change-flat', symbol: '→', delta: null };
  const previous = applicable[applicable.length - 2].percentage;
  const current = applicable[applicable.length - 1].percentage;
  if(typeof previous !== 'number' || typeof current !== 'number') return null;
  const delta = current - previous;
  if(Math.abs(delta) < 0.05) {
    return { label: 'No change since last return', shortLabel: 'No change', className: 'change-flat', symbol: '→', delta: 0 };
  }
  if(delta > 0) {
    return { label: `Up +${delta.toFixed(1)}pp since last return`, shortLabel: `+${delta.toFixed(1)}pp`, className: 'change-up', symbol: '↑', delta: delta };
  }
  return { label: `Down ${delta.toFixed(1)}pp since last return`, shortLabel: `${delta.toFixed(1)}pp`, className: 'change-down', symbol: '↓', delta: delta };
}
function hasHistoryAnomaly(history){
  const cleaned = normalizeTrendHistory(history);
  return cleaned.slice(-2).some((point) => typeof point.rawPercentage === 'number' && (point.rawPercentage < 0 || point.rawPercentage > 100));
}
function createChangeBadge(history, kpiName, compact=false, includeLastPrefix=false){
  const summary = getChangeSummary(history, kpiName);
  if(!summary) return '';
  let text = summary.label;
  if(compact){
    if(includeLastPrefix){
      if(summary.className === 'change-up' || summary.className === 'change-down'){
        text = `Last: ${summary.shortLabel}`;
      } else if(summary.className === 'change-missing'){
        text = 'Last: No return';
      } else if(summary.className === 'change-flat'){
        text = summary.delta === null ? 'Last: No prior' : 'Last: No change';
      } else {
        text = 'Last: N/A';
      }
    } else {
      text = summary.shortLabel;
    }
  }
  return `<span class="change-badge ${summary.className}" title="${summary.label}">${summary.symbol} ${text}</span>`;
}
function formatPercentCellCompressed(point, kpiName, fallbackText='No return'){
  if(!point) {
    if(String(fallbackText).toLowerCase().includes('prior')) return `<span class="cell-no-prior">No prior return ${createTooltip(kpiName, getNotApplicableText(kpiName))}</span>`;
    return `<span class="cell-no-return">${fallbackText}</span>`;
  }
  if(isImplicitNoReturnPoint(kpiName, point)) return '<span class="cell-no-return">No return</span>';
  if(point.applicable === false) return `N/A ${createTooltip(kpiName, getNotApplicableText(kpiName))}`;
  if(typeof point.percentage === 'number') return `${point.percentage.toFixed(1)}%`;
  return '<span class="cell-no-return">No return</span>';
}
function formatPercentCell(point, kpiName, fallbackText='No return'){
  if(!point) {
    if(String(fallbackText).toLowerCase().includes('prior')) return '<span class="no-prior">No previous applicable return</span>';
    return `<span class="no-return">${fallbackText}</span>`;
  }
  if(isImplicitNoReturnPoint(kpiName, point)) return '<span class="no-return">No return submitted</span>';
  if(point.applicable === false) return `N/A ${createTooltip(kpiName, getNotApplicableText(kpiName))}`;
  if(typeof point.percentage === 'number') return `${point.percentage.toFixed(1)}%`;
  return '<span class="no-return">No return submitted</span>';
}
function createCurrentPointFromKpi(kpi){
  if(!kpi) return null;
  return {
    date: 'Current',
    percentage: typeof kpi.percentage === 'number' ? kpi.percentage : null,
    rawPercentage: typeof kpi.percentage === 'number' ? kpi.percentage : null,
    number: typeof kpi.number === 'number' ? kpi.number : null,
    applicable: kpi.applicable !== false,
    exceeded: kpi.exceeded === true,
    completed_count: typeof kpi.completed_count === 'number' ? kpi.completed_count : null
  };
}
function getPreviousAndCurrentPoints(history,currentKpi=null){
  const cleaned = normalizeTrendHistory(history);
  if(cleaned.length === 0) return { previous: null, current: createCurrentPointFromKpi(currentKpi) };
  let current = cleaned[cleaned.length - 1];
  if(currentKpi && (!current || isImplicitNoReturnPoint('', current) || (current.applicable !== false && typeof current.percentage !== 'number'))){
    current = createCurrentPointFromKpi(currentKpi);
  }
  let previous = null;
  for(let i = cleaned.length - 2; i >= 0; i--){
    const candidate = cleaned[i];
    if(candidate && candidate.applicable && typeof candidate.percentage === 'number'){
      previous = candidate;
      break;
    }
  }
  return { previous, current };
}
function didSubmitReturnForPoint(kpiName, point){
  if(!point) return false;
  if(isImplicitNoReturnPoint(kpiName, point)) return false;
  if(point.applicable === false) return true;
  if(typeof point.percentage === 'number') return true;
  if(typeof point.rawPercentage === 'number') return true;
  if(point.number !== null && point.number !== undefined && !Number.isNaN(Number(point.number))) return true;
  return false;
}
function createPreviousReturnStatusBadge(history, kpiName){
  const cleaned = normalizeTrendHistory(history);
  if(cleaned.length < 2) return '<span class="change-badge change-flat">→ Prev: No prior period</span>';
  const previousPoint = cleaned[cleaned.length - 2];
  if(didSubmitReturnForPoint(kpiName, previousPoint)){
    return '<span class="change-badge change-prev-submitted" title="Previous reporting period had a submitted return">✓ Prev: Submitted</span>';
  }
  return '<span class="change-badge change-prev-missing" title="Previous reporting period did not have a submitted return">! Prev: No return</span>';
}
function getCompressedChangeCell(history, kpiName){
  const summary = getChangeSummary(history, kpiName);
  if(!summary) return '<span class="school-change na">N/A</span>';
  if(summary.className === 'change-missing') return '<span class="school-change down">No return</span>';
  if(summary.className === 'change-up') return `<span class="school-change up">↑ ${summary.delta.toFixed(1)}pp</span>`;
  if(summary.className === 'change-down') return `<span class="school-change down">↓ ${summary.delta.toFixed(1)}pp</span>`;
  if(summary.className === 'change-flat') {
    if(summary.delta === null) return '<span class="school-change flat">No previous return</span>';
    return '<span class="school-change flat">→ 0.0pp</span>';
  }
  return '<span class="school-change na">N/A</span>';
}
function createTrendMarkup(kpiData,kpiName){
  const facultyHistory = normalizeTrendHistory(kpiData.history);
  const universityHistory = normalizeTrendHistory(kpiData.university_history);
  if(facultyHistory.length < 2 && universityHistory.length < 2) return '';
  const safeId = `trend-${kpiName.replace(/[^a-zA-Z0-9]/g,'-')}`;
  const safeKpi = encodeURIComponent(kpiName);
  const facData = encodeURIComponent(JSON.stringify(facultyHistory));
  const uniData = encodeURIComponent(JSON.stringify(universityHistory));
  const primaryLabel = kpiData.is_university === true ? 'University' : DASHBOARD_CONFIG.entityLabel;
  const legend = universityHistory.length >= 2 && kpiData.is_university !== true
    ? `<div class="trend-legend"><span><i class="legend-line split"></i>${escapeHtml(primaryLabel)} (green above / red below university)</span><span><i class="legend-line dashed"></i>University</span></div>`
    : '';
  return `<div class="trend-wrap"><div class="trend-head"><div class="trend-title">Trend across reporting periods</div>${legend}</div><div class="trend-canvas-box"><canvas id="${safeId}" class="trend-canvas" data-kpi-name="${safeKpi}" data-primary-label="${escapeHtml(primaryLabel)}" data-faculty-history="${facData}" data-university-history="${uniData}"></canvas></div></div>`;
}
// Shorter badge wording for the narrow breakdown column; the full wording stays in the hover title.
// Colour a numeric percentage cell by its performance band.
function bandPercentCell(point, kpiName, cellHtml, variant){
  if(!point || point.applicable === false || isImplicitNoReturnPoint(kpiName, point) || typeof point.percentage !== 'number') return cellHtml;
  if(!/^-?\d+(\.\d+)?%$/.test(String(cellHtml).trim())) return cellHtml;
  const band = BAND_INFO[getPerformanceClass(point.percentage, true)];
  return `<span class="pct-chip pct-${variant} ${band.cls}" title="${band.name}: ${band.label}">${cellHtml}</span>`;
}
function shortenBreakdownBadge(badgeHtml){
  return String(badgeHtml || '')
    .replace('>○ N/A | No previous applicable return<', '>○ N/A<')
    .replace(/>○ N\/A \| Last applicable (-?\d+(?:\.\d+)?%)</, '>○ N/A · last $1<')
    .replace('>! Prev: No return<', '>! No return last period<')
    .replace('>✓ Prev: Submitted<', '>✓ Submitted last period<')
    .replace('>→ Prev: No prior period<', '>→ No prior period<');
}
function createSchoolTrendMarkup(history,kpiName,schoolName){
  const cleaned = normalizeTrendHistory(history);
  if(cleaned.length < 2) return `<div class="school-trend-inline"><div class="school-trend-empty">No prior period</div>${shortenBreakdownBadge(createPreviousReturnStatusBadge(cleaned,kpiName))}</div>`;
  const safeId = `school-trend-${schoolName.replace(/[^a-zA-Z0-9]/g,'-')}-${kpiName.replace(/[^a-zA-Z0-9]/g,'-')}`;
  const safeKpi = encodeURIComponent(kpiName);
  const historyData = encodeURIComponent(JSON.stringify(cleaned));
  const summary = getChangeSummary(cleaned, kpiName);
  const trendBadge = summary && summary.className === 'change-missing'
    ? createPreviousReturnStatusBadge(cleaned, kpiName)
    : createChangeBadge(cleaned, kpiName, true);
  return `<div class="school-trend-inline"><div class="spark-box"><canvas id="${safeId}" class="school-trend-canvas" data-kpi-name="${safeKpi}" data-school-history="${historyData}" aria-label="Trend for ${escapeHtml(schoolName)}"></canvas></div>${shortenBreakdownBadge(trendBadge)}</div>`;
}
const CHART_FONT = { family: "'Inter','Segoe UI',system-ui,-apple-system,sans-serif", size: 12 };
const CHART_TOOLTIP = {
  backgroundColor: '#0f172a',
  titleColor: '#ffffff',
  bodyColor: '#e2e8f0',
  titleFont: { ...CHART_FONT, weight: '600' },
  bodyFont: CHART_FONT,
  padding: 10,
  cornerRadius: 8,
  displayColors: true,
  boxWidth: 8,
  boxHeight: 8,
  usePointStyle: true
};
function renderTrendCharts(){
  if(typeof Chart === 'undefined') return;
  document.querySelectorAll('.trend-canvas').forEach((canvas) => {
    if(canvas.dataset.rendered === 'true') return;

    let facultyHistory = [];
    let universityHistory = [];
    try {
      facultyHistory = JSON.parse(decodeURIComponent(canvas.getAttribute('data-faculty-history') || '[]'));
    } catch (e) {
      facultyHistory = [];
    }
    try {
      universityHistory = JSON.parse(decodeURIComponent(canvas.getAttribute('data-university-history') || '[]'));
    } catch (e) {
      universityHistory = [];
    }

    const fac = normalizeTrendHistory(facultyHistory);
    const uni = normalizeTrendHistory(universityHistory);
    const kpiName = decodeURIComponent(canvas.getAttribute('data-kpi-name') || '');
    const primaryLabel = canvas.getAttribute('data-primary-label') || 'Faculty';
    const labels = (fac.length >= uni.length ? fac : uni).map((p) => p.date);
    if(labels.length < 2) {
      const wrap = canvas.closest('.trend-wrap');
      if(wrap) wrap.style.display = 'none';
      canvas.dataset.rendered = 'true';
      return;
    }

    const facMap = Object.fromEntries(fac.map((p) => [p.date, p.applicable ? p.percentage : null]));
    const uniMap = Object.fromEntries(uni.map((p) => [p.date, p.applicable ? p.percentage : null]));
    const facSeries = labels.map((l) => (facMap[l] ?? null));
    const uniSeries = labels.map((l) => (uniMap[l] ?? null));
    const bounds = buildTrendBounds([facSeries, uniSeries]);
    const showUniversity = uni.length >= 2 || !isUniversityDashboard;
    const ABOVE = BAND_COLORS.good;
    const BELOW = BAND_COLORS.poor;
    const BASE = BAND_COLORS.excellent;

    try {
      new Chart(canvas.getContext('2d'), {
      type: 'line',
      data: {
        labels: labels.map(formatReportDate),
        datasets: [
          {
            label: primaryLabel,
            data: facSeries,
            pointMeta: labels.map((label) => fac.find((point) => point.date === label) || { date: label, applicable: false, number: null }),
            borderColor: BASE,
            backgroundColor: BASE,
            pointBackgroundColor: '#ffffff',
            pointBorderColor: (ctx) => {
              const i = ctx.dataIndex;
              const f = facSeries[i];
              const u = uniSeries[i];
              if(!showUniversity || f === null || f === undefined || u === null || u === undefined) return BASE;
              return f >= u ? ABOVE : BELOW;
            },
            pointBorderWidth: 2,
            fill: false,
            tension: 0.25,
            borderWidth: 2.5,
            pointRadius: 4,
            pointHoverRadius: 6,
            segment: {
              borderColor: (ctx) => {
                const i = ctx.p0DataIndex;
                const f0 = facSeries[i];
                const f1 = facSeries[i + 1];
                const u0 = uniSeries[i];
                const u1 = uniSeries[i + 1];
                if(uni.length < 2 || [f0, f1, u0, u1].some((v) => v === null || v === undefined)) return BASE;
                const d0 = f0 - u0;
                const d1 = f1 - u1;
                if(d0 >= 0 && d1 >= 0) return ABOVE;
                if(d0 <= 0 && d1 <= 0) return BELOW;

                // Segment crosses the university line: split color at intersection.
                const p0 = ctx.p0;
                const p1 = ctx.p1;
                if(!p0 || !p1) return BASE;
                if(!Number.isFinite(p0.x) || !Number.isFinite(p0.y) || !Number.isFinite(p1.x) || !Number.isFinite(p1.y)) return BASE;
                const denom = d0 - d1;
                if(Math.abs(denom) < 1e-9) return BASE;
                const t = Math.max(0, Math.min(1, d0 / denom));
                if(!Number.isFinite(t) || t <= 0 || t >= 1) return d0 >= 0 ? ABOVE : BELOW;
                const grad = ctx.chart.ctx.createLinearGradient(p0.x, p0.y, p1.x, p1.y);
                const c0 = d0 > 0 ? ABOVE : BELOW;
                const c1 = d1 > 0 ? ABOVE : BELOW;
                grad.addColorStop(0, c0);
                grad.addColorStop(t, c0);
                grad.addColorStop(t, c1);
                grad.addColorStop(1, c1);
                return grad;
              }
            }
          },
          ...(showUniversity ? [{
            label: 'University',
            data: uniSeries,
            pointMeta: labels.map((label) => uni.find((point) => point.date === label) || { date: label, applicable: false, number: null }),
            borderColor: '#94a3b8',
            backgroundColor: '#94a3b8',
            borderDash: [5, 4],
            borderWidth: 2,
            fill: false,
            tension: 0.25,
            pointRadius: 3,
            pointHoverRadius: 5,
            pointBackgroundColor: '#94a3b8'
          }] : [])
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        devicePixelRatio: Math.max(2, window.devicePixelRatio || 1),
        interaction: { mode: 'index', intersect: false },
        layout: { padding: { top: 6, right: 8 } },
        plugins: {
          legend: { display: false },
          tooltip: {
            ...CHART_TOOLTIP,
            callbacks: {
              label: function(context) {
                const point = context.dataset.pointMeta?.[context.dataIndex];
                if(point && point.applicable === false) {
                  return `${context.dataset.label}: ${getNotApplicableText(kpiName)}`;
                }
                if(context.parsed.y === null || context.parsed.y === undefined) return `${context.dataset.label}: no data`;
                const countText = point && point.number !== null ? ` | In scope: ${Math.round(point.number)}` : '';
                return `${context.dataset.label}: ${context.parsed.y.toFixed(1)}%${countText}`;
              }
            }
          }
        },
        scales: {
          x: {
            ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 6, color: '#64748b', font: CHART_FONT },
            grid: { display: false },
            border: { color: '#cbd5e1' }
          },
          y: {
            // Round the axis to whole 5% steps so tick labels read cleanly.
            min: Math.floor(bounds.min / 5) * 5,
            max: Math.ceil(bounds.max / 5) * 5,
            ticks: { maxTicksLimit: 6, color: '#64748b', font: CHART_FONT, callback: function(v) { return (v < 0 || v > 100) ? '' : `${v}%`; } },
            grid: { color: '#eef1f5', drawBorder: false },
            border: { display: false }
          }
        }
      }
    });
    } catch (err) {
      console.error('Trend chart render failed', kpiName, err);
      const wrap = canvas.closest('.trend-wrap');
      if(wrap) wrap.style.display = 'none';
    }
    canvas.dataset.rendered = 'true';
  });
}
function renderSchoolTrendCharts(){
  if(typeof Chart === 'undefined') return;
  document.querySelectorAll('.school-trend-canvas').forEach((canvas) => {
    if(canvas.dataset.rendered === 'true') return;
    // Draw only once the breakdown is open; hidden canvases have no size.
    if(canvas.offsetWidth === 0) return;
    let history = [];
    try {
      history = JSON.parse(decodeURIComponent(canvas.getAttribute('data-school-history') || '[]'));
    } catch (e) {
      history = [];
    }
    const cleaned = normalizeTrendHistory(history);
    const kpiName = decodeURIComponent(canvas.getAttribute('data-kpi-name') || '');
    if(cleaned.length < 2) {
      canvas.dataset.rendered = 'true';
      return;
    }
    const series = cleaned.map((point) => (point.applicable ? point.percentage : null));
    const bounds = buildTrendBounds([series]);
    new Chart(canvas.getContext('2d'), {
      type: 'line',
      data: {
        labels: cleaned.map((point) => formatReportDate(point.date)),
        datasets: [{
          data: series,
          pointMeta: cleaned,
          borderColor: '#475569',
          backgroundColor: 'rgba(71,85,105,0.08)',
          fill: true,
          tension: 0.3,
          pointRadius: (ctx) => (ctx.dataIndex === series.length - 1 ? 2.5 : 0),
          pointBackgroundColor: '#0f172a',
          pointBorderWidth: 0,
          pointHitRadius: 8,
          borderWidth: 1.5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        devicePixelRatio: Math.max(2, window.devicePixelRatio || 1),
        layout: { padding: 3 },
        plugins: {
          legend: { display: false },
          tooltip: {
            ...CHART_TOOLTIP,
            displayColors: false,
            callbacks: {
              label: function(context) {
                const point = context.dataset.pointMeta?.[context.dataIndex];
                if(point && point.applicable === false) {
                  return getNotApplicableText(kpiName);
                }
                if(context.parsed.y === null || context.parsed.y === undefined) return 'No data';
                const countText = point && point.number !== null ? ` | In scope: ${Math.round(point.number)}` : '';
                return `${context.parsed.y.toFixed(1)}%${countText}`;
              }
            }
          }
        },
        scales: {
          x: { display: false },
          y: { display: false, min: bounds.min, max: bounds.max }
        },
        elements: {
          line: { capBezierPoints: true }
        }
      }
    });
    canvas.dataset.rendered = 'true';
  });
}
function getPerformanceClass(percentage,applicable=true){if(applicable===false)return 'performance-not-applicable';if(percentage===null||percentage===undefined)return 'performance-no-data';if(percentage>=95)return 'performance-excellent';if(percentage>=80)return 'performance-good';if(percentage>=60)return 'performance-warning';return 'performance-poor';}
function isMissingReturnValue(value){
  return value===null||value===undefined||(typeof value==='string'&&['/','-','–','—',''].includes(value.trim()));
}
function isImplicitNoReturnRecord(kpiName, percentage, number, applicable=true){
  if(applicable===false) return false;
  if(!kpiUsesDenominator(kpiName)) return false;
  const missingCount = number===null || number===undefined || Number.isNaN(Number(number));
  return missingCount && typeof Number(percentage)==='number' && !Number.isNaN(Number(percentage)) && Math.abs(Number(percentage)) < 1e-9;
}
function isMissingCurrentForKpi(kpiName, percentage, number, applicable=true){
  if(applicable===false) return false;
  return isMissingReturnValue(percentage) || isImplicitNoReturnRecord(kpiName, percentage, number, applicable);
}
function formatValue(percentage,number,applicable=true,kpiName=''){
  if(applicable===false) return getNotApplicableText(kpiName);
  if(isImplicitNoReturnRecord(kpiName, percentage, number, applicable)) return '<span class="no-return">No return submitted</span>';
  if(isMissingReturnValue(percentage)) return '<span class="no-return">No return submitted</span>';
  const zeroScoreText = getZeroScoreText(kpiName);
  if(zeroScoreText && Math.abs(Number(percentage)) < 1e-9) return zeroScoreText;
  if(number===null||number===undefined) return `${percentage.toFixed(1)}%`;
  return `${percentage.toFixed(1)}%`;
}
function formatDisplayTextWithCount(kpiName,percentage,number,applicable=true,exceeded=false,completedCount=null){
  if(exceeded){
    const label=kpiMeta[kpiName]&&kpiMeta[kpiName].denominator_label?kpiMeta[kpiName].denominator_label:'items';
    if(number !== null && number !== undefined && Math.abs(Number(number)) < 1e-9){
      if(completedCount !== null){
        return `<span class="exceeded-badge">Activity reported (0 ${label} scheduled, ${completedCount} carried out)</span>`;
      }
      return `<span class="exceeded-badge">Activity reported (0 ${label} scheduled)</span>`;
    }
    if(completedCount !== null){
      return `${percentage.toFixed(1)}% (${completedCount} carried out, ${Math.round(number)} ${label} scheduled)`;
    }
    return `${percentage.toFixed(1)}% (${Math.round(number)} ${label} in scope)`;
  }
  if(applicable===false) return getNotApplicableText(kpiName);
  if(isImplicitNoReturnRecord(kpiName, percentage, number, applicable)) return '<span class="no-return">No return submitted</span>';
  if(isMissingReturnValue(percentage)) return '<span class="no-return">No return submitted</span>';
  const zeroScoreText = getZeroScoreText(kpiName);
  if(zeroScoreText && Math.abs(Number(percentage)) < 1e-9) return zeroScoreText;
  if(number===null||number===undefined) return `${percentage.toFixed(1)}%`;
  const label=kpiMeta[kpiName]&&kpiMeta[kpiName].denominator_label?kpiMeta[kpiName].denominator_label:'records';
  return `${percentage.toFixed(1)}% (${Math.round(number)} ${label} in scope)`;
}
function createUniversityComparison(kpiName,kpiData){
  const compact = currentViewMode === 'compressed';
  const changeBadge = createChangeBadge(kpiData.history, kpiName, compact, false);
  const changeHtml = changeBadge ? `<span class="col-change">${changeBadge}</span>` : '<span class="col-change"></span>';
  if(kpiData.is_university === true) return changeHtml;
  const facultyPercentage = kpiData.percentage;
  const universityPercentage = kpiData.university_percentage;
  const unavailable = `<span class="col-uni"><span class="delta unavailable" title="University comparison unavailable this period">${compact ? 'Unavailable' : 'University comparison unavailable'}</span></span>`;
  if(kpiData.applicable===false||kpiData.university_applicable===false||facultyPercentage===null||universityPercentage===null){
    return changeBadge ? unavailable + changeHtml : '<span class="col-uni"></span>' + changeHtml;
  }
  const absDelta = Math.abs(facultyPercentage - universityPercentage).toFixed(1);
  const uniText = compact ? `(uni ${universityPercentage.toFixed(1)}%)` : `vs university ${universityPercentage.toFixed(1)}%`;
  let pill = '';
  switch(kpiData.comparison){
    case 'above':
      pill = `<span class="delta above" title="Above university by +${absDelta}pp">▲ +${absDelta}pp <span class="delta-label">${uniText}</span></span>`;
      break;
    case 'below':
      pill = `<span class="delta below" title="Below university by -${absDelta}pp">▼ −${absDelta}pp <span class="delta-label">${uniText}</span></span>`;
      break;
    case 'equal':
      pill = `<span class="delta equal" title="Equal to university">= <span class="delta-label">Equal to university (${universityPercentage.toFixed(1)}%)</span></span>`;
      break;
    default:
      return changeBadge ? unavailable + changeHtml : '<span class="col-uni"></span>' + changeHtml;
  }
  return `<span class="col-uni">${pill}</span>${changeHtml}`;
}
function getUniversityMarker(kpiData){
  if(kpiData.is_university === true) return null;
  if(kpiData.applicable===false||kpiData.university_applicable===false) return null;
  if(typeof kpiData.percentage !== 'number' || typeof kpiData.university_percentage !== 'number') return null;
  if(!['above','below','equal'].includes(kpiData.comparison)) return null;
  return kpiData.university_percentage;
}
// Split the existing display text into a headline value and its supporting detail.
function buildHeroMarkup(kpiName,kpiData,performanceClass){
  const display = formatDisplayTextWithCount(kpiName, kpiData.percentage, kpiData.number, kpiData.applicable, kpiData.exceeded, kpiData.completed_count);
  const match = /^(-?\d+(?:\.\d+)?%)(?: \((.*)\))?$/.exec(String(display).trim());
  if(match){
    const band = BAND_INFO[performanceClass] || BAND_INFO['performance-no-data'];
    const chip = `<span class="band-chip ${band.cls}" title="${band.name}: ${band.label}">${band.name}</span>`;
    const sub = match[2] ? `<div class="kpi-hero-sub">${match[2].charAt(0).toUpperCase() + match[2].slice(1)}</div>` : '';
    return { numeric: true, html: `<div class="kpi-hero"><span class="kpi-hero-value num">${match[1]}</span>${chip}</div>${sub}` };
  }
  let stateClass = '';
  if(display.includes('no-return')) stateClass = 'state-missing';
  else if(display.includes('exceeded-badge')) stateClass = 'state-exceeded';
  return { numeric: false, html: `<div class="kpi-state ${stateClass}">${display}</div>` };
}
function buildMeterMarkup(kpiData,performanceClass){
  if(typeof kpiData.percentage !== 'number' || kpiData.applicable === false) return '';
  const band = BAND_INFO[performanceClass];
  if(!band || band.cls === 'band-na') return '';
  const width = Math.max(0, Math.min(100, kpiData.percentage));
  const uni = getUniversityMarker(kpiData);
  const marker = uni === null ? '' : `<div class="kpi-meter-marker" style="left:${Math.max(0, Math.min(100, uni))}%" data-label="Uni" title="University: ${uni.toFixed(1)}%"></div>`;
  return `<div class="kpi-meter ${band.cls}" role="img" aria-label="${kpiData.percentage.toFixed(1)} percent"><div class="kpi-meter-fill" style="width:${width}%"></div>${marker}</div>`;
}
function pluralise(count, singular, plural){
  return `${count} ${count === 1 ? singular : plural}`;
}
function toggleCard(cardElement,kpiName){
  cardElement.classList.toggle('expanded');
  const button = cardElement.querySelector('.expand-btn');
  if(button) button.setAttribute('aria-expanded', cardElement.classList.contains('expanded') ? 'true' : 'false');
  if(cardElement.classList.contains('expanded')) setTimeout(renderSchoolTrendCharts, 0);
}
function refreshDashboard(){
  initializeFacultyDashboard();
  setTimeout(()=>{positionTooltips();applyCompressedValueTitles();renderTrendCharts();renderSchoolTrendCharts();},50);
}
function sortFacultyCards(order){
  currentSortOrder=order;
  document.getElementById('sortHighLow').classList.toggle('active',order==='highLow');
  document.getElementById('sortLowHigh').classList.toggle('active',order==='lowHigh');
  document.getElementById('sortHighLow').setAttribute('aria-pressed',order==='highLow');
  document.getElementById('sortLowHigh').setAttribute('aria-pressed',order==='lowHigh');
  refreshDashboard();
}
function toggleViewMode(mode){
  currentViewMode=mode;
  document.getElementById('detailedView').classList.toggle('active',mode==='detailed');
  document.getElementById('compressedView').classList.toggle('active',mode==='compressed');
  document.getElementById('detailedView').setAttribute('aria-pressed',mode==='detailed');
  document.getElementById('compressedView').setAttribute('aria-pressed',mode==='compressed');
  const container=document.querySelector('.faculty-overview');
  const instructions=document.getElementById('instructions');
  if(mode==='compressed'){
    container.classList.add('compressed-view');
    instructions.textContent=DASHBOARD_CONFIG.compactHint;
  } else {
    container.classList.remove('compressed-view');
    instructions.textContent=DASHBOARD_CONFIG.detailedHint;
  }
  refreshDashboard();
}
function renderHeaderMeta(){
  const meta = document.getElementById('headerMeta');
  if(!meta || !facultyData.kpis) return;
  let dates = [];
  Object.values(facultyData.kpis).forEach((kpi) => {
    const history = normalizeTrendHistory(kpi.history);
    if(history.length > dates.length) dates = history.map((point) => point.date);
  });
  const parts = [];
  if(dates.length > 0) parts.push(`<span>Latest return: <strong>${escapeHtml(formatReportDate(dates[dates.length - 1]))}</strong></span>`);
  if(dates.length > 1) parts.push(`<span>Trend covers <strong>${dates.length} reporting periods</strong> from ${escapeHtml(formatReportDate(dates[0]))}</span>`);
  parts.push(`<span><strong>${Object.keys(facultyData.kpis).length}</strong> KPIs</span>`);
  meta.innerHTML = parts.join('');
}
// The band a KPI is counted in on the summary tiles (missing returns are unrated).
function getKpiSummaryBand(kpiName, kpiData){
  const missing = isMissingCurrentForKpi(kpiName, kpiData.percentage, kpiData.number, kpiData.applicable);
  return missing ? 'performance-no-data' : getPerformanceClass(kpiData.percentage, kpiData.applicable);
}
function toggleBandFilter(bandClass){
  if(activeBandFilters.has(bandClass)) activeBandFilters.delete(bandClass); else activeBandFilters.add(bandClass);
  syncBandFilterUi();
  refreshDashboard();
}
function clearBandFilters(){
  activeBandFilters.clear();
  syncBandFilterUi();
  refreshDashboard();
}
function syncBandFilterUi(){
  const filtering = activeBandFilters.size > 0;
  document.querySelectorAll('.band-tile[data-band]').forEach((tile) => {
    const selected = activeBandFilters.has(tile.dataset.band);
    tile.classList.toggle('selected', selected);
    tile.setAttribute('aria-pressed', selected ? 'true' : 'false');
  });
  const tiles = document.querySelector('.band-tiles');
  if(tiles) tiles.classList.toggle('filtering', filtering);
  let bar = document.getElementById('filterBar');
  if(!bar){
    bar = document.createElement('div');
    bar.id = 'filterBar';
    bar.className = 'filter-bar';
    bar.setAttribute('role', 'status');
    const grid = document.getElementById('facultyKpis');
    const head = document.querySelector('.compact-head');
    grid.parentNode.insertBefore(bar, head || grid);
  }
  if(!filtering){ bar.hidden = true; bar.innerHTML = ''; return; }
  const total = Object.keys(facultyData.kpis || {}).length;
  const shown = Object.entries(facultyData.kpis || {}).filter(([n, k]) => activeBandFilters.has(getKpiSummaryBand(n, k))).length;
  const chips = [...activeBandFilters].map((cls) => {
    const band = BAND_INFO[cls];
    return `<button type="button" class="band-chip ${band.cls} filter-chip" onclick="toggleBandFilter('${cls}')" aria-label="Remove filter: ${band.name}">${band.name}<span aria-hidden="true">×</span></button>`;
  }).join('');
  bar.hidden = false;
  bar.innerHTML = `<span class="filter-label">Showing <strong>${shown}</strong> of ${total} KPIs</span>${chips}<button type="button" class="filter-clear" onclick="clearBandFilters()">Show all KPIs</button>`;
}
function renderSummary(){
  const host = document.getElementById('glance');
  if(!host || !facultyData.kpis) return;
  const kpis = Object.entries(facultyData.kpis);
  const counts = { 'performance-excellent':0, 'performance-good':0, 'performance-warning':0, 'performance-poor':0 };
  let unrated = 0, improved = 0, declined = 0, above = 0, compared = 0;
  kpis.forEach(([kpiName, kpiData]) => {
    const cls = getKpiSummaryBand(kpiName, kpiData);
    if(cls in counts) counts[cls] += 1; else unrated += 1;
    const change = getChangeSummary(kpiData.history, kpiName);
    if(change && change.className === 'change-up') improved += 1;
    if(change && change.className === 'change-down') declined += 1;
    if(!isUniversityDashboard && ['above','below','equal'].includes(kpiData.comparison)){
      compared += 1;
      if(kpiData.comparison === 'above') above += 1;
    }
  });
  const rated = kpis.length - unrated;
  const tiles = Object.entries(counts).map(([cls, count]) => {
    const band = BAND_INFO[cls];
    return `<button type="button" class="band-tile ${band.cls}" data-band="${cls}" aria-pressed="false" ${count === 0 ? 'disabled' : ''} onclick="toggleBandFilter('${cls}')" title="${count === 0 ? 'No KPIs in this band' : `Show only ${band.name} KPIs`}"><div class="band-tile-count num">${count}<small>KPI${count === 1 ? '' : 's'}</small></div><div class="band-tile-label">${band.name}</div><div class="band-tile-range">${band.label}</div></button>`;
  }).join('');
  const distribution = rated > 0
    ? Object.entries(counts).filter(([, c]) => c > 0).map(([cls, c]) => `<span class="${BAND_INFO[cls].cls}" style="flex:${c}" title="${BAND_INFO[cls].label}: ${c}"></span>`).join('')
    : '';
  const unratedNote = unrated > 0 ? `${pluralise(unrated, 'KPI is', 'KPIs are')} not rated this period (not applicable or no return).` : 'All KPIs have a rated value this period.';

  const units = facultyData.schools ? Object.entries(facultyData.schools) : [];
  const noReturnUnits = units.filter(([, unit]) => {
    const entries = Object.entries(unit.kpis || {});
    return entries.length > 0 && entries.every(([kpiName, k]) => isMissingCurrentForKpi(kpiName, k.percentage, k.number, k.applicable));
  }).map(([name]) => name);
  const missingKpiReturns = units.reduce((total, [, unit]) => total + Object.entries(unit.kpis || {}).filter(([kpiName, k]) => isMissingCurrentForKpi(kpiName, k.percentage, k.number, k.applicable)).length, 0);

  let rows = `<div class="stat-row"><span class="stat-label"><span class="stat-icon up">↑</span>Improved since last return</span><span class="stat-value num">${improved}</span></div>`;
  rows += `<div class="stat-row"><span class="stat-label"><span class="stat-icon down">↓</span>Declined since last return</span><span class="stat-value num">${declined}</span></div>`;
  if(!isUniversityDashboard){
    rows += `<div class="stat-row"><span class="stat-label"><span class="stat-icon above">▲</span>Above university average</span><span class="stat-value num">${compared > 0 ? `${above}<small> of ${compared}</small>` : '<small>Not available</small>'}</span></div>`;
  }
  rows += `<div class="stat-row${noReturnUnits.length > 0 ? ' has-sub' : ''}"><span class="stat-label"><span class="stat-icon missing">!</span>${escapeHtml(DASHBOARD_CONFIG.unitTitle)} with no return</span><span class="stat-value num">${noReturnUnits.length}<small> of ${units.length}</small></span></div>`;
  if(noReturnUnits.length > 0) rows += `<div class="stat-sub">${noReturnUnits.map((name) => `<span class="unit-chip">${escapeHtml(name)}</span>`).join('')}</div>`;

  host.innerHTML = `
    <div class="panel glance-bands">
      <div class="panel-title"><span>Current performance</span><span class="panel-title-note">${rated} of ${kpis.length} KPIs rated · select a band to filter</span></div>
      <div class="band-tiles">${tiles}</div>
      <div class="band-distribution" aria-hidden="true">${distribution}</div>
      <div class="glance-note">${unratedNote}${missingKpiReturns > 0 ? ` ${pluralise(missingKpiReturns, 'KPI return is', 'KPI returns are')} missing across ${escapeHtml(DASHBOARD_CONFIG.unitPlural)}.` : ''}</div>
    </div>
    <div class="panel glance-stats">
      <div class="panel-title"><span>At a glance</span></div>
      <div class="stat-list">${rows}</div>
    </div>`;
}
function initializeFacultyDashboard(){
  applyExceededInference();
  const facultyKpisContainer = document.getElementById('facultyKpis');
  facultyKpisContainer.innerHTML = '';
  if(!facultyData.kpis) return;

  let kpiEntries = Object.entries(facultyData.kpis);
  if(currentSortOrder === 'highLow') {
    kpiEntries.sort((a, b) => {
      const aPercentage = a[1].applicable === false || a[1].percentage === null ? -1 : a[1].percentage;
      const bPercentage = b[1].applicable === false || b[1].percentage === null ? -1 : b[1].percentage;
      return bPercentage - aPercentage;
    });
  } else {
    kpiEntries.sort((a, b) => {
      const aPercentage = a[1].applicable === false || a[1].percentage === null ? 999 : a[1].percentage;
      const bPercentage = b[1].applicable === false || b[1].percentage === null ? 999 : b[1].percentage;
      return aPercentage - bPercentage;
    });
  }

  if(activeBandFilters.size > 0){
    kpiEntries = kpiEntries.filter(([kpiName, kpiData]) => activeBandFilters.has(getKpiSummaryBand(kpiName, kpiData)));
  }

  for(const [kpiName, kpiData] of kpiEntries){
    const card = document.createElement('div');
    card.className = 'kpi-card';
    let cardHasMissingReturns = isMissingCurrentForKpi(kpiName, kpiData.percentage, kpiData.number, kpiData.applicable);
    let missingReturnsCount = 0;
    if(cardHasMissingReturns) card.classList.add('no-return-current');
    card.id = `card-${kpiName.replace(/[^a-zA-Z0-9]/g, '-')}`;
    const performanceClass = getPerformanceClass(kpiData.percentage, kpiData.applicable);
    const cardBand = BAND_INFO[performanceClass] || BAND_INFO['performance-no-data'];
    card.classList.add(cardBand.cls);
    let schoolBreakdownHtml = '';
    let schoolCount = 0;

    const hasSchools = facultyData.schools && Object.keys(facultyData.schools).length > 0;
    if(hasSchools){
      let schoolEntries = Object.entries(facultyData.schools).filter(([schoolName, schoolData]) => schoolData.kpis && schoolData.kpis[kpiName]);
      missingReturnsCount = schoolEntries.reduce((count, [, schoolData]) => {
        const schoolKpi = schoolData.kpis[kpiName];
        return count + (isMissingCurrentForKpi(kpiName, schoolKpi.percentage, schoolKpi.number, schoolKpi.applicable) ? 1 : 0);
      }, 0);
      if(missingReturnsCount > 0) cardHasMissingReturns = true;

      if(currentSortOrder === 'highLow') {
        schoolEntries.sort((a, b) => {
          const aPoint = a[1].kpis[kpiName];
          const bPoint = b[1].kpis[kpiName];
          const aPerc = aPoint.applicable === false || aPoint.percentage === null ? -1 : aPoint.percentage;
          const bPerc = bPoint.applicable === false || bPoint.percentage === null ? -1 : bPoint.percentage;
          return bPerc - aPerc;
        });
      } else {
        schoolEntries.sort((a, b) => {
          const aPoint = a[1].kpis[kpiName];
          const bPoint = b[1].kpis[kpiName];
          const aPerc = aPoint.applicable === false || aPoint.percentage === null ? 999 : aPoint.percentage;
          const bPerc = bPoint.applicable === false || bPoint.percentage === null ? 999 : bPoint.percentage;
          return aPerc - bPerc;
        });
      }

      schoolCount = schoolEntries.length;
      if(schoolEntries.length > 0){
        const breakdownTitle = `Breakdown by ${DASHBOARD_CONFIG.unitSingular} (${schoolEntries.length})`;
        schoolBreakdownHtml = `<div class="school-breakdown-content"><div class="school-breakdown-title">${breakdownTitle}</div>`;

        if(currentViewMode === 'compressed'){
          schoolBreakdownHtml += `<div class="compressed-school-table"><div class="compressed-school-row header"><div>${escapeHtml(DASHBOARD_CONFIG.unitColumn)}</div><div>In scope</div><div>Previous</div><div>Current</div><div>Change</div></div>`;
          for(const [schoolName, schoolData] of schoolEntries){
            const schoolKpi = schoolData.kpis[kpiName];
            const number = schoolKpi.number;
            const countValue = (number !== null && number !== undefined && number !== '/' && !isNaN(parseFloat(number))) ? parseFloat(number) : null;
            const points = getPreviousAndCurrentPoints(schoolKpi.history, isUniversityDashboard ? schoolKpi : null);
            if(isMissingCurrentForKpi(kpiName, schoolKpi.percentage, schoolKpi.number, schoolKpi.applicable)) cardHasMissingReturns = true;
            const schoolWarning = createOutOfRangeBadge(schoolKpi.percentage, 'school-warning');
            const hasNoScope = countValue === null;
            const isExceeded = schoolKpi.exceeded === true;
            const hasNoCurrent = !isExceeded && (!points.current || isImplicitNoReturnPoint(kpiName, points.current) || (points.current.applicable !== false && typeof points.current.percentage !== 'number'));
            if(hasNoScope && hasNoCurrent){
              schoolBreakdownHtml += `<div class="compressed-school-row no-return-row"><div class="school-name">${schoolName}${schoolWarning}</div><div class="row-no-return-msg">No return submitted for this period</div></div>`;
            } else {
              const scopeText = getInScopeTextCompressed(kpiName, countValue, schoolKpi);
              const previousText = bandPercentCell(points.previous, kpiName, formatPercentCellCompressed(points.previous, kpiName, 'No prior applicable return'), 'prev');
              const currentText = bandPercentCell(points.current, kpiName, formatPercentCellCompressed(points.current, kpiName, 'No return'), 'current');
              const changeText = getCompressedChangeCell(schoolKpi.history, kpiName);
              const rowBand = (points.current && points.current.applicable !== false && typeof points.current.percentage === 'number' && !isImplicitNoReturnPoint(kpiName, points.current))
                ? BAND_INFO[getPerformanceClass(points.current.percentage, true)].cls : 'band-na';
              schoolBreakdownHtml += `<div class="compressed-school-row ${rowBand}"><div class="school-name">${schoolName}${schoolWarning}</div><div class="school-scope">${scopeText}</div><div class="school-prev num">${previousText}</div><div class="school-current num">${currentText}</div><div>${changeText}</div></div>`;
            }
          }
          schoolBreakdownHtml += '</div>';
        } else {
          schoolBreakdownHtml += '<div class="school-list">';
          for(const [schoolName, schoolData] of schoolEntries){
            const schoolKpi = schoolData.kpis[kpiName];
            const percentage = schoolKpi.percentage;
            const number = schoolKpi.number;
            const applicable = schoolKpi.applicable !== false;
            let barWidth = 0;
            let barColor = BAND_COLORS.neutral;
            let labelText = '';
            let isNumericLabel = false;
            let stateClass = '';

            const hasNumber = (number !== null && number !== undefined && number !== '/' && !isNaN(parseFloat(number)));
            const countValue = hasNumber ? parseFloat(number) : null;
            const hasZeroCount = countValue === 0;
            const zeroScoreText = getZeroScoreText(kpiName);
            const missingCurrent = isMissingCurrentForKpi(kpiName, percentage, number, applicable);

            if(missingCurrent){
              cardHasMissingReturns = true;
              labelText = '<span class="no-return">No return submitted</span>';
            } else if(!applicable){
              labelText = getNotApplicableText(kpiName);
            } else if(percentage === null || percentage === undefined || (typeof percentage === 'string' && percentage.trim() === '/')) {
              labelText = '<span class="no-return">No return submitted</span>';
            } else {
              const numericPercentage = parseFloat(percentage);
              if(!isNaN(numericPercentage)) {
                if(zeroScoreText && Math.abs(numericPercentage) < 1e-9) {
                  labelText = zeroScoreText;
                } else {
                  const boundedPercentage = Math.max(0, Math.min(100, numericPercentage));
                  if(hasZeroCount && !schoolKpi.exceeded) {
                    labelText = getNotApplicableText(kpiName);
                  } else if(hasZeroCount && schoolKpi.exceeded) {
                    const excLabel = kpiMeta[kpiName] && kpiMeta[kpiName].denominator_label ? kpiMeta[kpiName].denominator_label : 'items';
                    labelText = schoolKpi.completed_count !== null
                      ? `Activity reported (0 ${excLabel} scheduled, ${schoolKpi.completed_count} carried out)`
                      : `Activity reported (0 ${excLabel} scheduled)`;
                    stateClass = 'state-exceeded';
                  } else {
                    barWidth = boundedPercentage;
                    const schoolPerfClass = getPerformanceClass(numericPercentage, true);
                    switch(schoolPerfClass){
                      case 'performance-excellent': barColor = BAND_COLORS.excellent; break;
                      case 'performance-good': barColor = BAND_COLORS.good; break;
                      case 'performance-warning': barColor = BAND_COLORS.warning; break;
                      case 'performance-poor': barColor = BAND_COLORS.poor; break;
                      default: barColor = BAND_COLORS.neutral;
                    }
                    labelText = `${numericPercentage.toFixed(1)}%`;
                    isNumericLabel = true;
                  }
                }
              } else {
                labelText = '<span class="no-return">No return submitted</span>';
              }
            }

            if(labelText === '') labelText = '<span class="no-return">No return submitted</span>';

            const schoolLabel = (kpiMeta[kpiName] && kpiMeta[kpiName].denominator_label) ? kpiMeta[kpiName].denominator_label : 'records';
            const countDisplay = countValue !== null ? `<div class="school-count">${Math.round(countValue)} ${schoolLabel} in scope</div>` : '';
            const schoolTrendHtml = createSchoolTrendMarkup(schoolKpi.history, kpiName, schoolName);
            const measureHtml = isNumericLabel
              ? `<div class="bar-track"><div class="bar-fill" style="width:${barWidth}%;background:${barColor};"></div></div><span class="bar-value num">${labelText}</span>`
              : `<span class="school-state ${stateClass}">${labelText}</span>`;

            const schoolWarning = createOutOfRangeBadge(schoolKpi.percentage, 'school-warning');
            schoolBreakdownHtml += `<div class="school-row"><div><div class="school-name">${schoolName}${schoolWarning}</div>${countDisplay}</div><div class="school-measure">${measureHtml}</div><div class="school-trend">${schoolTrendHtml}</div></div>`;
          }
          schoolBreakdownHtml += '</div>';
        }

        schoolBreakdownHtml += '</div>';
      }
    }

    if(missingReturnsCount === 0 && cardHasMissingReturns) missingReturnsCount = 1;

    const trendHtml = currentViewMode === 'compressed' ? '' : createTrendMarkup(kpiData, kpiName);
    const hasExpandableContent = schoolBreakdownHtml !== '';
    const warningBadgeHtml = createOutOfRangeBadge(kpiData.percentage);
    const missingCardBadge = cardHasMissingReturns ? `<span class="missing-return-flag" title="${escapeHtml(DASHBOARD_CONFIG.unitTitle)} with no return for this KPI this period">${pluralise(missingReturnsCount, 'missing return', 'missing returns')}</span>` : '';
    const comparisonHtml = createUniversityComparison(kpiName, kpiData);
    const questionText = (kpiMeta[kpiName] && kpiMeta[kpiName].question) ? kpiMeta[kpiName].question : '';
    const hero = buildHeroMarkup(kpiName, kpiData, performanceClass);
    const meterHtml = buildMeterMarkup(kpiData, performanceClass);
    const unitWord = schoolCount === 1 ? DASHBOARD_CONFIG.unitSingular : DASHBOARD_CONFIG.unitPlural;
    const footerHtml = hasExpandableContent
      ? `<div class="card-footer"><button type="button" class="expand-btn" aria-expanded="false" aria-label="Show or hide ${escapeHtml(DASHBOARD_CONFIG.unitSingular)} breakdown for ${escapeHtml(kpiName)}"><span class="expand-text"><span class="when-closed">Show ${schoolCount} ${escapeHtml(unitWord)}</span><span class="when-open">Hide breakdown</span></span><span class="chev" aria-hidden="true">▾</span></button></div>`
      : (currentViewMode === 'compressed' ? '<div class="card-footer"></div>' : '<div class="card-footer-spacer"></div>');
    const flags = `<span class="kpi-flags">${warningBadgeHtml}${missingCardBadge}</span>`;

    if(currentViewMode === 'compressed') {
      card.innerHTML = `<div class="kpi-title-wrap"><span class="kpi-title">${kpiName}</span>${createKpiInfoButton(kpiName, true)}${flags}</div><div class="kpi-hero-wrap">${hero.html}</div>${comparisonHtml}${footerHtml}${schoolBreakdownHtml}`;
    } else {
      card.innerHTML = `<div class="kpi-head"><div class="kpi-title-wrap"><h3 class="kpi-title">${kpiName}</h3>${flags}</div>${createKpiInfoButton(kpiName)}</div><div class="kpi-question">${questionText}</div><div class="kpi-body"><div class="kpi-hero-wrap">${hero.html}</div><div class="kpi-deltas">${comparisonHtml}</div></div>${meterHtml}${trendHtml}${footerHtml}${schoolBreakdownHtml}`;
    }
    if(hasExpandableContent) {
      card.onclick = (event) => {
        if(event.target.closest('.info-btn')) return;
        if(event.target.closest('.school-breakdown-content') && !event.target.closest('.expand-btn')) return;
        toggleCard(card, kpiName);
      };
    } else {
      card.style.cursor = 'default';
      card.onclick = null;
    }

    facultyKpisContainer.appendChild(card);
  }
}
document.addEventListener('keydown', (event) => { if(event.key === 'Escape') closeInfoDialog(); });
renderHeaderMeta();
renderSummary();
refreshDashboard();
window.addEventListener('resize', () => { setTimeout(() => { renderTrendCharts(); renderSchoolTrendCharts(); }, 100); });
