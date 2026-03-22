const I18N = {
  zh: {
    documentTitle: 'SPX 稀有信号分析器',
    loading: '正在加载区间数据...',
    heroEyebrow: '标普 500 稀有信号监控',
    heroTitle: 'SPX K 线、稀有信号与后续统计。',
    heroText: '切换日线、周线、月线，直接看当前值得关注的 rare signals、历史出现次数、后续上涨概率，以及相对无条件基准的收益分布。',
    summaryLabel: '当前解读',
    dataFreshness: '数据更新至',
    chartKicker: 'K 线图',
    chartHeading: '价格结构',
    eventKicker: '当前稀有信号',
    eventHeading: '实时信号板',
    detailKicker: '信号详情',
    definitionHeading: '信号定义',
    definitionSubtitle: '当前触发条件与对应阈值',
    occurrenceHeading: '最近几次发生',
    distributionHeading: '收益分布',
    distributionSubtitle: '条件分布 vs 无条件基准',
    historicalExamplesHeading: '历史样本',
    historyDateHead: '日期',
    historyReturnHead: '当根涨跌',
    historyRangeHead: '振幅',
    historyDownHead: '连跌计数',
    historyBearishHead: '阴线计数',
    emptyState: '当前区间没有满足阈值的 rare signal。你仍然可以切换 interval，或者继续观察下一根 K 线是否触发新的极端模式。',
    noLiveRareEvent: '当前没有实时稀有信号',
    noMethodology: '暂无方法说明',
    latest: '最新',
    showing: '显示',
    of: '/',
    bars: '根',
    liveNow: '实时触发',
    watchlist: '观察中',
    historicalHits: '历史出现',
    rarityScore: '稀有分',
    conditionalLabel: '条件样本',
    unconditionalLabel: '无条件样本',
    occurrenceSignalReturn: '信号当根涨跌',
    recentHitsSuffix: '次最近命中',
    intervals: {
      daily: '日线',
      weekly: '周线',
      monthly: '月线',
    },
    ranges: {
      '1Y': '1年',
      '3Y': '3年',
      '5Y': '5年',
      MAX: '全部',
    },
    cards: {
      conditionalUpProb: '条件上涨概率',
      unconditionalUpProb: '无条件上涨概率',
      conditionalMean: '条件均值',
      unconditionalMean: '无条件均值',
      conditionalMedian: '条件中位数',
      unconditionalMedian: '无条件中位数',
      currentVolumeRegime: '当前量能分位',
      currentVixRegime: '当前 VIX 分位',
    },
  },
  en: {
    documentTitle: 'SPX Rare Event Analyzer',
    loading: 'Loading interval data...',
    heroEyebrow: 'S&P 500 rare event monitor',
    heroTitle: 'SPX candles, rare signals, and forward stats.',
    heroText: 'Switch across daily, weekly, and monthly bars to view live rare signals, historical hit counts, forward win rates, and return distributions against the unconditional baseline.',
    summaryLabel: 'Current read',
    dataFreshness: 'Data updated through',
    chartKicker: 'Candlestick',
    chartHeading: 'Price structure',
    eventKicker: 'Current rare events',
    eventHeading: 'Live signal board',
    detailKicker: 'Event detail',
    definitionHeading: 'Signal definition',
    definitionSubtitle: 'Exact trigger rules and active thresholds',
    occurrenceHeading: 'Recent occurrences',
    distributionHeading: 'Return distribution',
    distributionSubtitle: 'Conditional vs unconditional baseline',
    historicalExamplesHeading: 'Historical examples',
    historyDateHead: 'Date',
    historyReturnHead: 'Close return',
    historyRangeHead: 'Range',
    historyDownHead: 'Down streak',
    historyBearishHead: 'Bearish streak',
    emptyState: 'No rare signal is active for this interval right now. Switch the interval or wait for the next bar to see whether a new setup is triggered.',
    noLiveRareEvent: 'No live rare event',
    noMethodology: 'No methodology note',
    latest: 'Latest',
    showing: 'showing',
    of: 'of',
    bars: 'bars',
    liveNow: 'Live now',
    watchlist: 'Watchlist',
    historicalHits: 'historical hits',
    rarityScore: 'rarity score',
    conditionalLabel: 'Conditional',
    unconditionalLabel: 'Unconditional',
    occurrenceSignalReturn: 'Signal return',
    recentHitsSuffix: 'most recent hits',
    intervals: {
      daily: 'Daily',
      weekly: 'Weekly',
      monthly: 'Monthly',
    },
    ranges: {
      '1Y': '1Y',
      '3Y': '3Y',
      '5Y': '5Y',
      MAX: 'Max',
    },
    cards: {
      conditionalUpProb: 'Conditional Up Prob',
      unconditionalUpProb: 'Unconditional Up Prob',
      conditionalMean: 'Conditional Mean',
      unconditionalMean: 'Unconditional Mean',
      conditionalMedian: 'Conditional Median',
      unconditionalMedian: 'Unconditional Median',
      currentVolumeRegime: 'Current Volume Regime',
      currentVixRegime: 'Current VIX Regime',
    },
  },
};

const state = {
  interval: 'daily',
  language: localStorage.getItem('spx-language') || 'zh',
  selectedEventId: null,
  selectedWindow: null,
  selectedRange: '1Y',
  chart: null,
  candleSeries: null,
  volumeSeries: null,
  histogram: null,
  eventMap: {},
  currentBars: [],
  currentDashboard: null,
};

function dict() {
  return I18N[state.language];
}

function intervalLabel(interval) {
  return dict().intervals[interval] || interval;
}

function rangeLabel(range) {
  return dict().ranges[range] || range;
}

function naLabel() {
  return state.language === 'zh' ? '暂无' : 'N/A';
}

function percent(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return naLabel();
  }
  return `${(value * 100).toFixed(1)}%`;
}

function number(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return naLabel();
  }
  return `${value.toFixed(2)}%`;
}

function compactNumber(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return naLabel();
  }
  return `${value.toFixed(1)}%`;
}

function percentileLabel(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return naLabel();
  }
  return `P${Math.round(value * 100)}`;
}

function eventTitle(event) {
  return state.language === 'zh' ? event.title : (event.title_en || event.title);
}

function eventDescription(event) {
  return state.language === 'zh'
    ? event.plain_language_description
    : (event.plain_language_description_en || event.plain_language_description);
}

function eventMethodology(event) {
  return state.language === 'zh'
    ? event.methodology_note
    : (event.methodology_note_en || event.methodology_note);
}

function dashboardSummary(dashboard) {
  if (!dashboard) return '';
  return state.language === 'zh' ? dashboard.summary : (dashboard.summary_en || dashboard.summary);
}

async function getJSON(path) {
  const response = await fetch(path);
  if (!response.ok) {
    throw new Error(`Request failed: ${path}`);
  }
  return response.json();
}

function setLoading(isLoading) {
  const el = document.getElementById('loadingState');
  if (!el) return;
  el.textContent = dict().loading;
  el.classList.toggle('visible', isLoading);
}

function applyStaticLanguage() {
  const copy = dict();
  document.title = copy.documentTitle;
  document.documentElement.lang = state.language === 'zh' ? 'zh-CN' : 'en';

  document.getElementById('loadingState').textContent = copy.loading;
  document.getElementById('heroEyebrow').textContent = copy.heroEyebrow;
  document.getElementById('heroTitle').textContent = copy.heroTitle;
  document.getElementById('heroText').textContent = copy.heroText;
  document.getElementById('summaryLabel').textContent = copy.summaryLabel;
  document.getElementById('chartKicker').textContent = copy.chartKicker;
  document.getElementById('chartHeading').textContent = copy.chartHeading;
  document.getElementById('eventKicker').textContent = copy.eventKicker;
  document.getElementById('eventHeading').textContent = copy.eventHeading;
  document.getElementById('detailKicker').textContent = copy.detailKicker;
  document.getElementById('definitionHeading').textContent = copy.definitionHeading;
  document.getElementById('definitionSubtitle').textContent = copy.definitionSubtitle;
  document.getElementById('occurrenceHeading').textContent = copy.occurrenceHeading;
  document.getElementById('distributionHeading').textContent = copy.distributionHeading;
  document.getElementById('distributionSubtitle').textContent = copy.distributionSubtitle;
  document.getElementById('historicalExamplesHeading').textContent = copy.historicalExamplesHeading;
  document.getElementById('historyDateHead').textContent = copy.historyDateHead;
  document.getElementById('historyReturnHead').textContent = copy.historyReturnHead;
  document.getElementById('historyRangeHead').textContent = copy.historyRangeHead;
  document.getElementById('historyDownHead').textContent = copy.historyDownHead;
  document.getElementById('historyBearishHead').textContent = copy.historyBearishHead;

  const dataFreshness = document.getElementById('dataFreshness');
  dataFreshness.textContent = `${copy.dataFreshness} ${dataFreshness.dataset.date}`;

  document.querySelectorAll('.interval-button').forEach((button) => {
    button.textContent = intervalLabel(button.dataset.interval);
  });
  document.querySelectorAll('.range-button').forEach((button) => {
    button.textContent = rangeLabel(button.dataset.range);
  });
  document.querySelectorAll('.language-button').forEach((button) => {
    button.classList.toggle('active', button.dataset.lang === state.language);
  });
}

function setupChart() {
  const container = document.getElementById('candles');
  state.chart = LightweightCharts.createChart(container, {
    layout: {
      background: { color: '#08111d' },
      textColor: '#c5d3e6',
    },
    grid: {
      vertLines: { color: 'rgba(131, 152, 181, 0.12)' },
      horzLines: { color: 'rgba(131, 152, 181, 0.12)' },
    },
    width: container.clientWidth,
    height: container.clientHeight,
    rightPriceScale: {
      borderColor: 'rgba(131, 152, 181, 0.16)',
      scaleMargins: { top: 0.06, bottom: 0.24 },
    },
    timeScale: {
      borderColor: 'rgba(131, 152, 181, 0.16)',
    },
    crosshair: {
      vertLine: { color: 'rgba(247, 184, 1, 0.4)' },
      horzLine: { color: 'rgba(247, 184, 1, 0.4)' },
    },
  });

  state.candleSeries = state.chart.addCandlestickSeries({
    upColor: '#13c296',
    downColor: '#f45b69',
    borderVisible: false,
    wickUpColor: '#13c296',
    wickDownColor: '#f45b69',
  });

  state.volumeSeries = state.chart.addHistogramSeries({
    priceFormat: { type: 'volume' },
    priceScaleId: 'volume',
  });
  state.volumeSeries.priceScale().applyOptions({
    scaleMargins: { top: 0.80, bottom: 0.0 },
    visible: false,
  });

  window.addEventListener('resize', () => {
    state.chart.applyOptions({ width: container.clientWidth, height: container.clientHeight });
  });
}

function applyChartRange() {
  if (!state.currentBars.length) return;
  if (state.selectedRange === 'MAX') {
    state.chart.timeScale().fitContent();
    return;
  }

  const latest = new Date(`${state.currentBars[state.currentBars.length - 1].date}T00:00:00Z`);
  const years = Number(state.selectedRange.replace('Y', '')) || 1;
  const from = new Date(latest);
  from.setFullYear(from.getFullYear() - years);

  state.chart.timeScale().setVisibleRange({
    from: Math.floor(from.getTime() / 1000),
    to: Math.floor(latest.getTime() / 1000),
  });
}

function renderSeries(seriesPayload) {
  state.currentBars = seriesPayload.bars.map((bar) => ({
    ...bar,
    timestamp: Math.floor(new Date(`${bar.date}T00:00:00Z`).getTime() / 1000),
  }));

  state.candleSeries.setData(
    state.currentBars.map((bar) => ({
      time: bar.timestamp,
      open: bar.open,
      high: bar.high,
      low: bar.low,
      close: bar.close,
    }))
  );

  state.volumeSeries.setData(
    state.currentBars
      .filter((bar) => bar.volume !== null && bar.volume !== undefined)
      .map((bar) => ({
        time: bar.timestamp,
        value: bar.volume,
        color: bar.close >= bar.open ? 'rgba(19, 194, 150, 0.45)' : 'rgba(244, 91, 105, 0.45)',
      }))
  );

  applyChartRange();
  const copy = dict();
  document.getElementById('chartMeta').textContent =
    `${copy.latest} ${intervalLabel(state.interval)}: ${seriesPayload.latest_date} · ${copy.showing} ${seriesPayload.chart_bar_count} ${copy.of} ${seriesPayload.full_bar_count} ${copy.bars}`;
}

function renderSummary() {
  document.getElementById('summaryText').textContent = dashboardSummary(state.currentDashboard) || dict().loading;
}

function renderEmptyDetail() {
  document.getElementById('detailTitle').textContent = dict().noLiveRareEvent;
  document.getElementById('detailTag').textContent = dict().noMethodology;
  document.getElementById('statsGrid').innerHTML = '';
  document.getElementById('definitionList').innerHTML = '';
  document.getElementById('examplesBody').innerHTML = '';
  document.getElementById('occurrenceBody').innerHTML = '';
  document.getElementById('occurrenceHeadRow').innerHTML = '';
  document.getElementById('windowSwitch').innerHTML = '';
  document.getElementById('occurrenceSubtitle').textContent = '';
  if (state.histogram) {
    state.histogram.destroy();
    state.histogram = null;
  }
}

function renderEvents(eventsPayload) {
  const copy = dict();
  const list = document.getElementById('eventList');
  list.innerHTML = '';
  state.eventMap = {};

  if (!eventsPayload.events.length) {
    state.selectedEventId = null;
    list.innerHTML = `<div class="empty-state">${copy.emptyState}</div>`;
    renderEmptyDetail();
    return;
  }

  const preferredEventId = state.selectedEventId;
  eventsPayload.events.forEach((event, index) => {
    state.eventMap[event.event_id] = event;
    const button = document.createElement('button');
    button.type = 'button';
    const selected = preferredEventId ? preferredEventId === event.event_id : index === 0;
    button.className = `event-card${selected ? ' selected' : ''}`;
    button.dataset.eventId = event.event_id;
    button.innerHTML = `
      <div class="event-rank">0${index + 1}</div>
      <div class="event-main">
        <div class="event-title-row">
          <h3>${eventTitle(event)}</h3>
          <span class="event-badge">${event.current_triggered ? copy.liveNow : copy.watchlist}</span>
        </div>
        <p>${eventDescription(event)}</p>
        <div class="event-foot">
          <span>${event.historical_count} ${copy.historicalHits}</span>
          <span>${copy.rarityScore} ${event.rarity_score}</span>
        </div>
      </div>
    `;
    button.addEventListener('click', async () => {
      document.querySelectorAll('.event-card').forEach((item) => item.classList.remove('selected'));
      button.classList.add('selected');
      state.selectedEventId = event.event_id;
      await loadEventDetail();
    });
    list.appendChild(button);
  });

  state.selectedEventId = state.eventMap[preferredEventId]
    ? preferredEventId
    : eventsPayload.events[0].event_id;
}

function renderStats(event) {
  const copy = dict();
  document.getElementById('detailTitle').textContent = eventDescription(event);
  document.getElementById('detailTag').textContent = eventMethodology(event) || copy.noMethodology;

  const firstWindow = event.forward_windows[0];
  if (!state.selectedWindow || !event.forward_windows.includes(state.selectedWindow)) {
    state.selectedWindow = firstWindow;
  }

  const switchWrap = document.getElementById('windowSwitch');
  switchWrap.innerHTML = '';
  event.forward_windows.forEach((windowKey) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `window-button${windowKey === state.selectedWindow ? ' active' : ''}`;
    button.textContent = windowKey;
    button.addEventListener('click', () => {
      state.selectedWindow = windowKey;
      renderStats(event);
      renderDistribution(event);
    });
    switchWrap.appendChild(button);
  });

  const conditional = event.conditional_stats[state.selectedWindow];
  const unconditional = event.unconditional_stats[state.selectedWindow];
  const cards = [
    [copy.cards.conditionalUpProb, `${percent(conditional.up_probability)} / n=${conditional.sample_size}`],
    [copy.cards.unconditionalUpProb, `${percent(unconditional.up_probability)} / n=${unconditional.sample_size}`],
    [copy.cards.conditionalMean, number(conditional.mean_return)],
    [copy.cards.unconditionalMean, number(unconditional.mean_return)],
    [copy.cards.conditionalMedian, number(conditional.median_return)],
    [copy.cards.unconditionalMedian, number(unconditional.median_return)],
    [copy.cards.currentVolumeRegime, percentileLabel(event.current_context.volume_percentile)],
    [copy.cards.currentVixRegime, percentileLabel(event.current_context.vix_level_percentile)],
  ];

  document.getElementById('statsGrid').innerHTML = cards.map(([label, value]) => `
    <div class="stat-card">
      <div class="stat-label">${label}</div>
      <div class="stat-value">${value}</div>
    </div>
  `).join('');

  document.getElementById('definitionList').innerHTML = (state.language === 'zh' ? event.rule_lines : (event.rule_lines_en || event.rule_lines || []))
    .map((line) => `<div class="definition-item">${line}</div>`)
    .join('');

  document.getElementById('occurrenceHeadRow').innerHTML = [
    `<th>${copy.historyDateHead}</th>`,
    `<th>${copy.occurrenceSignalReturn}</th>`,
    ...event.forward_windows.map((windowKey) => `<th>${windowKey}</th>`),
  ].join('');

  document.getElementById('occurrenceBody').innerHTML = event.recent_occurrences.map((occurrence) => `
    <tr>
      <td>${occurrence.date}</td>
      <td>${compactNumber(occurrence.close_return)}</td>
      ${event.forward_windows.map((windowKey) => `<td>${compactNumber(occurrence.forward_returns[windowKey])}</td>`).join('')}
    </tr>
  `).join('');

  document.getElementById('occurrenceSubtitle').textContent =
    state.language === 'zh'
      ? `最近 ${event.recent_occurrences.length} 次命中`
      : `${event.recent_occurrences.length} ${copy.recentHitsSuffix}`;

  document.getElementById('examplesBody').innerHTML = event.historical_examples.map((example) => `
    <tr>
      <td>${example.date}</td>
      <td>${number(example.close_return)}</td>
      <td>${number(example.range_pct)}</td>
      <td>${example.down_streak}</td>
      <td>${example.bearish_streak}</td>
    </tr>
  `).join('');
}

function renderDistribution(event) {
  const copy = dict();
  const bins = event.distribution_bins[state.selectedWindow];
  const conditional = bins.conditional;
  const unconditional = bins.unconditional;
  const labels = conditional.map((bin) => `${((bin.start + bin.end) / 2).toFixed(1)}%`);
  const datasets = [
    {
      label: copy.conditionalLabel,
      data: conditional.map((bin) => bin.count),
      borderColor: '#f7b801',
      backgroundColor: 'rgba(247, 184, 1, 0.12)',
      borderWidth: 2,
      tension: 0.28,
      fill: true,
      pointRadius: 0,
    },
    {
      label: copy.unconditionalLabel,
      data: unconditional.map((bin) => bin.count),
      borderColor: '#5bc0eb',
      backgroundColor: 'rgba(91, 192, 235, 0.10)',
      borderWidth: 2,
      tension: 0.28,
      fill: true,
      pointRadius: 0,
    },
  ];

  const ctx = document.getElementById('distributionChart');
  if (state.histogram) {
    state.histogram.destroy();
  }
  state.histogram = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'index',
        intersect: false,
      },
      scales: {
        x: {
          ticks: { color: '#9fb2c8', maxRotation: 0, autoSkip: true },
          grid: { display: false },
        },
        y: {
          ticks: { color: '#9fb2c8' },
          grid: { color: 'rgba(131, 152, 181, 0.12)' },
        },
      },
      plugins: {
        legend: {
          labels: { color: '#c5d3e6' },
        },
        tooltip: {
          backgroundColor: 'rgba(7, 17, 29, 0.96)',
          borderColor: 'rgba(148, 170, 197, 0.18)',
          borderWidth: 1,
        },
      },
    },
  });
}

async function loadEventDetail() {
  if (!state.selectedEventId) return;
  const event = state.eventMap[state.selectedEventId];
  if (!event) return;
  renderStats(event);
  renderDistribution(event);
}

function rerenderForLanguage() {
  applyStaticLanguage();
  renderSummary();
  if (state.currentDashboard) {
    renderSeries({
      ...state.currentDashboard.series,
      interval: state.currentDashboard.interval,
    });
    renderEvents({
      events: state.currentDashboard.events,
      interval: state.currentDashboard.interval,
    });
    loadEventDetail();
  } else {
    renderEmptyDetail();
  }
}

async function loadInterval(interval) {
  state.interval = interval;
  setLoading(true);
  try {
    const dashboard = await getJSON(`/api/dashboard?interval=${interval}`);
    state.currentDashboard = dashboard;
    const dataFreshness = document.getElementById('dataFreshness');
    dataFreshness.dataset.date = dashboard.data_freshness.latest_merged_date;
    dataFreshness.textContent = `${dict().dataFreshness} ${dashboard.data_freshness.latest_merged_date}`;
    renderSeries({
      ...dashboard.series,
      interval: dashboard.interval,
    });
    renderEvents({
      events: dashboard.events,
      interval: dashboard.interval,
    });
    renderSummary();
    await loadEventDetail();
  } finally {
    setLoading(false);
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  applyStaticLanguage();
  setupChart();

  document.querySelectorAll('.interval-button').forEach((button) => {
    button.addEventListener('click', async () => {
      document.querySelectorAll('.interval-button').forEach((item) => item.classList.remove('active'));
      button.classList.add('active');
      await loadInterval(button.dataset.interval);
    });
  });

  document.querySelectorAll('.range-button').forEach((button) => {
    button.addEventListener('click', () => {
      document.querySelectorAll('.range-button').forEach((item) => item.classList.remove('active'));
      button.classList.add('active');
      state.selectedRange = button.dataset.range;
      applyChartRange();
    });
  });

  document.querySelectorAll('.language-button').forEach((button) => {
    button.addEventListener('click', () => {
      state.language = button.dataset.lang;
      localStorage.setItem('spx-language', state.language);
      rerenderForLanguage();
    });
  });

  await loadInterval('daily');
});
