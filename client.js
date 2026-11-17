/**
 * dsh-personal-finance-workbench · Client 侧入口（已构建单文件）
 *
 * 个人财务管理工作台：整页业务面板 + 按需 AI 对话。
 * 流程：数据接入 → 数据体检 → 口径映射 → 分析结果 → 交付文件。
 */

window.__ModuleLoader__.load({
  id: 'dsh-personal-finance-workbench',
  factory: (require) => {
    var React = require('react');
    var h = React.createElement;
    var useState = React.useState;
    var useEffect = React.useEffect;
    var useMemo = React.useMemo;

    var API = '/api/personal-finance-workbench';

    /** apply() 注入的宿主上下文，供面板调用原生能力（目录选择器）。 */
    var host = null;

    var STEPS = [
      { key: 'data', label: '数据接入', hint: '选择 Excel / CSV 流水' },
      { key: 'quality', label: '数据体检', hint: '看字段与问题' },
      { key: 'mapping', label: '口径映射', hint: '确认五件事' },
      { key: 'analysis', label: '分析结果', hint: '收支 / 结构 / 异常' },
      { key: 'deliver', label: '交付文件', hint: 'xlsx + docx' },
    ];

    var MAPPING_FIELDS = [
      { key: 'date', label: '日期列', required: false, hint: '用于月度收支，缺失时合并为「未标注日期」' },
      { key: 'amount', label: '金额列', required: true, hint: '收入或支出的金额（正数）' },
      { key: 'direction', label: '收支方向', required: false, hint: '收入 / 支出列；没有时按金额正负判断' },
      { key: 'category', label: '分类列', required: false, hint: '餐饮 / 房租 / 工资 / 理财' },
      { key: 'account', label: '账户列', required: false, hint: '微信 / 支付宝 / 银行卡' },
    ];

    async function post(path, body) {
      var response = await fetch(API + path, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body || {}),
      });
      var data = await response.json().catch(function () { return {}; });
      if (!response.ok) throw new Error(data.error || ('请求失败（HTTP ' + response.status + '）'));
      return data;
    }

    async function getInfo() {
      var response = await fetch(API + '/info', { cache: 'no-store' });
      var data = await response.json().catch(function () { return {}; });
      if (!response.ok) throw new Error(data.error || '读取工作台状态失败');
      return data;
    }

    function money(value) {
      if (value == null || !isFinite(value)) return '—';
      return Number(value).toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function pct(value, digits) {
      if (value == null || !isFinite(value)) return '—';
      return (value * 100).toFixed(digits == null ? 1 : digits) + '%';
    }

    function pctDelta(value) {
      if (value == null || !isFinite(value)) return '—';
      return (value >= 0 ? '+' : '') + (value * 100).toFixed(1) + '%';
    }

    function bytesText(bytes) {
      if (!bytes) return '0 KB';
      if (bytes < 1024 * 1024) return Math.max(1, Math.round(bytes / 1024)) + ' KB';
      return (bytes / 1024 / 1024).toFixed(1) + ' MB';
    }

    function shortPath(p) {
      if (!p) return '';
      return p.length > 68 ? '…' + p.slice(-66) : p;
    }

    var CHART_COLORS = ['#2563eb', '#f59e0b', '#10b981', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#84cc16', '#f97316', '#64748b'];

    function compactNumber(value) {
      if (value == null || !isFinite(value)) return '—';
      var abs = Math.abs(value);
      if (abs >= 100000000) return (value / 100000000).toFixed(1) + '亿';
      if (abs >= 10000) return (value / 10000).toFixed(abs >= 1000000 ? 0 : 1) + '万';
      if (abs >= 1000) return Math.round(value).toLocaleString('zh-CN');
      return (Math.round(value * 10) / 10).toString();
    }

    function niceMax(value) {
      var v = Math.abs(value || 0);
      if (!v || !isFinite(v)) return 1;
      var exp = Math.pow(10, Math.floor(Math.log10(v)));
      var norm = v / exp;
      var step = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
      return step * exp;
    }

    function niceStep(value) {
      var v = Math.abs(value || 0);
      if (!v || !isFinite(v)) return 1;
      var exp = Math.pow(10, Math.floor(Math.log10(v)));
      var norm = v / exp;
      var step = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
      return step * exp;
    }

    function axisUnit(maxAbs) {
      if (maxAbs >= 100000000) return 'yi';
      if (maxAbs >= 10000) return 'wan';
      return 'num';
    }

    function formatAxis(value, maxAbs) {
      if (value === 0) return '0';
      var abs = Math.abs(value);
      var unit = axisUnit(maxAbs || abs);
      if (unit === 'yi') return (value / 100000000).toFixed(1) + '亿';
      if (unit === 'wan') {
        var text = (value / 10000).toFixed(1);
        return (text.slice(-2) === '.0' ? text.slice(0, -2) : text) + '万';
      }
      if (abs >= 1000) return String(Math.round(value));
      return String(Math.round(value * 100) / 100);
    }

    function axisPct(value) {
      return (Math.round(value * 1000) / 10) + '%';
    }

    function ChartPanel(props) {
      return h('div', { className: 'finPanel' + (props.wide ? ' finPanelWide' : '') },
        h('div', { className: 'finPanelHead' },
          h('b', null, props.title),
          props.hint ? h('span', null, props.hint) : null),
        props.children);
    }

    function ChartLegend(props) {
      return h('div', { className: 'finLegend' }, (props.items || []).map(function (item, index) {
        return h('span', { key: index }, h('i', { style: { background: item.color } }), item.name);
      }));
    }

    function GroupedBars(props) {
      var rows = props.rows || [];
      var series = props.series || [];
      if (!rows.length) return h('div', { className: 'finEmpty' }, '暂无可绘制的数据');
      var wide = !!props.wide;
      var W = wide ? 960 : 470;
      var H = wide ? 264 : 232;
      var padL = wide ? 52 : 44;
      var padR = 14;
      var padT = 14;
      var padB = wide ? 36 : 52;
      var plotW = W - padL - padR;
      var plotH = H - padT - padB;
      var hi = 0;
      var lo = 0;
      rows.forEach(function (row) {
        (row.values || []).forEach(function (value) {
          if (!isFinite(value)) return;
          if (value > hi) hi = value;
          if (value < lo) lo = value;
        });
      });
      if (hi <= 0 && lo >= 0) hi = 1;
      var span = hi - lo;
      if (!isFinite(span) || span <= 0) span = 1;
      var axisStep = niceStep(span / 4);
      var top = Math.ceil(hi / axisStep) * axisStep;
      var bottom = Math.floor(lo / axisStep) * axisStep;
      var range = top - bottom;
      if (!isFinite(range) || range <= 0) { top = span; bottom = 0; range = span; }
      var tickCount = Math.round(range / axisStep);
      if (!isFinite(tickCount) || tickCount < 2) tickCount = 4;
      var yOf = function (value) { return padT + plotH - ((value - bottom) / range) * plotH; };
      var zeroY = yOf(0);
      var maxAbs = Math.max(Math.abs(top), Math.abs(bottom));
      var groupW = plotW / rows.length;
      var inner = groupW * 0.72;
      var barW = Math.max(2, inner / Math.max(1, series.length));
      var labelEvery = rows.length > 10 && !wide ? 2 : 1;
      var parts = [];
      var t;
      for (t = 0; t <= tickCount; t += 1) {
        var tickValue = bottom + axisStep * t;
        var gy = yOf(tickValue);
        var isZero = Math.abs(tickValue) < axisStep / 1000;
        parts.push(h('line', { key: 'g' + t, x1: padL, x2: W - padR, y1: gy, y2: gy, stroke: 'currentColor', strokeOpacity: isZero ? 0.3 : 0.12, strokeDasharray: isZero ? null : '3 4' }));
        parts.push(h('text', { key: 'gt' + t, x: padL - 8, y: gy + 4, textAnchor: 'end' }, formatAxis(tickValue, maxAbs)));
      }
      rows.forEach(function (row, rowIndex) {
        var baseX = padL + rowIndex * groupW + (groupW - inner) / 2;
        (row.values || []).forEach(function (value, sIndex) {
          var sv = isFinite(value) ? value : 0;
          var magnitude = Math.abs(sv) === 0 ? 0 : Math.max(1.5, (Math.abs(sv) / range) * plotH);
          var by = sv >= 0 ? zeroY - magnitude : zeroY;
          var fill = props.colorBySign ? (sv >= 0 ? '#10b981' : '#ef4444') : (series[sIndex] ? series[sIndex].color : CHART_COLORS[sIndex % CHART_COLORS.length]);
          parts.push(h('rect', {
            key: 'b' + rowIndex + '-' + sIndex,
            x: baseX + sIndex * barW + 0.8,
            y: by,
            width: Math.max(1.6, barW - 1.6),
            height: magnitude,
            rx: 2,
            fill: fill,
            opacity: 0.92,
          }, h('title', null, (row.label || '') + ' · ' + (series[sIndex] ? series[sIndex].name + ' ' : '') + (props.format ? props.format(sv) : String(sv)))));
        });
        if (rowIndex % labelEvery === 0) {
          var lx = padL + rowIndex * groupW + groupW / 2;
          var ly = H - padB + (wide ? 16 : 12);
          parts.push(h('text', {
            key: 'l' + rowIndex,
            x: lx,
            y: ly,
            textAnchor: 'end',
            transform: 'rotate(-32 ' + lx + ' ' + ly + ')',
          }, String(row.label || '')));
        }
      });
      return h('div', null,
        h('svg', { className: 'finChart', viewBox: '0 0 ' + W + ' ' + H, preserveAspectRatio: 'xMidYMid meet' }, parts),
        series.length > 1 ? h(ChartLegend, { items: series }) : null);
    }

    function TrendLines(props) {
      var rows = props.rows || [];
      var series = props.series || [];
      if (!rows.length) return h('div', { className: 'finEmpty' }, '暂无可绘制的数据');
      var wide = !!props.wide;
      var W = wide ? 960 : 470;
      var H = wide ? 264 : 232;
      var padL = wide ? 54 : 44;
      var padR = 14;
      var padT = 16;
      var padB = wide ? 36 : 52;
      var plotW = W - padL - padR;
      var plotH = H - padT - padB;
      var peak = 0;
      var low = 0;
      rows.forEach(function (row) {
        (row.values || []).forEach(function (value) {
          if (!isFinite(value)) return;
          if (value > peak) peak = value;
          if (value < low) low = value;
        });
      });
      var dataMin = props.zeroBase === false ? low : 0;
      var dataMax = peak;
      var axisStep = niceStep((dataMax - dataMin) / 4 || 1);
      var yMin = Math.floor(dataMin / axisStep) * axisStep;
      var yMax = Math.ceil(dataMax / axisStep) * axisStep;
      if (yMax <= yMin) yMax = yMin + axisStep;
      var span = yMax - yMin;
      var tickCount = Math.max(2, Math.min(7, Math.round(span / axisStep)));
      var yFor = function (value) { return padT + plotH - ((value - yMin) / span) * plotH; };
      var parts = [];
      var t;
      for (t = 0; t <= tickCount; t += 1) {
        var tickValue = yMin + axisStep * t;
        var gy = yFor(tickValue);
        parts.push(h('line', { key: 'g' + t, x1: padL, x2: W - padR, y1: gy, y2: gy, stroke: 'currentColor', strokeOpacity: t === 0 ? 0.28 : 0.12, strokeDasharray: t === 0 ? null : '3 4' }));
        parts.push(h('text', { key: 'gt' + t, x: padL - 8, y: gy + 4, textAnchor: 'end' }, props.pctAxis ? axisPct(tickValue) : formatAxis(tickValue, Math.max(Math.abs(yMin), Math.abs(yMax)))));
      }
      var labelEvery = rows.length > 10 && !wide ? 2 : 1;
      rows.forEach(function (row, rowIndex) {
        if (rowIndex % labelEvery !== 0) return;
        var lx = padL + (plotW / (rows.length - 1 || 1)) * rowIndex;
        var ly = H - padB + (wide ? 16 : 12);
        parts.push(h('text', {
          key: 'l' + rowIndex,
          x: lx,
          y: ly,
          textAnchor: 'end',
          transform: 'rotate(-32 ' + lx + ' ' + ly + ')',
        }, String(row.label || '')));
      });
      series.forEach(function (line, sIndex) {
        var points = [];
        rows.forEach(function (row, rowIndex) {
          var value = (row.values || [])[sIndex];
          if (!isFinite(value)) return;
          var x = padL + (plotW / (rows.length - 1 || 1)) * rowIndex;
          points.push([x, yFor(value)]);
        });
        if (points.length > 1 && props.area !== false) {
          var area = 'M' + points[0][0].toFixed(1) + ',' + (padT + plotH).toFixed(1);
          points.forEach(function (pt) { area += ' L' + pt[0].toFixed(1) + ',' + pt[1].toFixed(1); });
          area += ' L' + points[points.length - 1][0].toFixed(1) + ',' + (padT + plotH).toFixed(1) + ' Z';
          parts.push(h('path', { key: 'area' + sIndex, d: area, fill: line.color, opacity: 0.10 }));
        }
        if (points.length > 1) {
          parts.push(h('polyline', {
            key: 'line' + sIndex,
            points: points.map(function (pt) { return pt[0].toFixed(1) + ',' + pt[1].toFixed(1); }).join(' '),
            fill: 'none',
            stroke: line.color,
            strokeWidth: 2.2,
            strokeLinejoin: 'round',
            strokeLinecap: 'round',
          }));
        }
        points.forEach(function (pt, pointIndex) {
          var value = (rows[pointIndex].values || [])[sIndex];
          parts.push(h('circle', { key: 'd' + sIndex + '-' + pointIndex, cx: pt[0], cy: pt[1], r: 2.6, fill: line.color },
            h('title', null, (rows[pointIndex].label || '') + ' · ' + line.name + ' ' + (props.format ? props.format(value) : String(value)))));
        });
      });
      return h('div', null,
        h('svg', { className: 'finChart', viewBox: '0 0 ' + W + ' ' + H, preserveAspectRatio: 'xMidYMid meet' }, parts),
        series.length > 1 ? h(ChartLegend, { items: series }) : null);
    }

    function DonutChart(props) {
      var items = (props.items || []).filter(function (item) { return item.value > 0; });
      var colors = props.colors || CHART_COLORS;
      if (!items.length) return h('div', { className: 'finEmpty' }, '暂无可绘制的数据');
      var size = props.size || 170;
      var thickness = props.thickness || 22;
      var radius = (size - thickness) / 2;
      var circumference = 2 * Math.PI * radius;
      var center = size / 2;
      var total = 0;
      items.forEach(function (item) { total += item.value; });
      if (!total) total = 1;
      var offset = 0;
      var arcs = items.map(function (item, index) {
        var length = (item.value / total) * circumference;
        var node = h('circle', {
          key: 'arc' + index,
          cx: center,
          cy: center,
          r: radius,
          fill: 'none',
          stroke: colors[index % colors.length],
          strokeWidth: thickness,
          strokeDasharray: length.toFixed(2) + ' ' + (circumference - length).toFixed(2),
          strokeDashoffset: (-offset).toFixed(2),
          transform: 'rotate(-90 ' + center + ' ' + center + ')',
        }, h('title', null, item.name + ' · ' + (props.format ? props.format(item.value) : String(item.value)) + ' · ' + ((item.value / total) * 100).toFixed(1) + '%'));
        offset += length;
        return node;
      });
      return h('div', { className: 'finDonut' },
        h('svg', { width: size, height: size, viewBox: '0 0 ' + size + ' ' + size, style: { flex: 'none' } },
          arcs,
          h('text', { x: center, y: center - 2, textAnchor: 'middle', style: { fontSize: 17, fontWeight: 650, fill: 'var(--fin-text,#111827)' } }, props.centerValue || ''),
          h('text', { x: center, y: center + 16, textAnchor: 'middle' }, props.centerLabel || '')),
        h('div', { className: 'finLegendCol' }, items.slice(0, props.maxItems || 8).map(function (item, index) {
          return h('div', { key: item.name, className: 'finLegendRow' },
            h('i', { style: { background: colors[index % colors.length] } }),
            h('span', { className: 'finLegendName', title: item.name }, item.name),
            h('em', null, props.format ? props.format(item.value) : String(item.value)),
            h('b', null, ((item.value / total) * 100).toFixed(1) + '%'));
        })));
    }

    function HBars(props) {
      var items = props.items || [];
      if (!items.length) return h('div', { className: 'finEmpty' }, '暂无可绘制的数据');
      var peak = 1;
      items.forEach(function (item) { if (Math.abs(item.value || 0) > peak) peak = Math.abs(item.value || 0); });
      return h('div', null, items.map(function (item, index) {
        var width = Math.max(1.5, (Math.abs(item.value || 0) / peak) * 100);
        return h('div', { key: (item.name || index) + '', className: 'finHBarRow' },
          h('span', { className: 'finHBarName', title: item.name }, item.name),
          h('div', { className: 'finHBarTrack' }, h('i', { style: { width: width + '%', background: item.color || props.color || '#2563eb' } })),
          h('b', null, props.format ? props.format(item.value) : String(item.value)));
      }));
    }

    var CSS = [
      '.finRoot{box-sizing:border-box;display:flex;flex-direction:column;width:100%;height:100%;min-height:0;background:#f5f7fb;color:#0f172a;font-family:var(--fin-font,"Inter","SF Pro Text",-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif);font-size:13px;overflow:hidden;color-scheme:light;position:relative;isolation:isolate}',
      '.finRoot::before{content:"";position:absolute;inset:0;z-index:-1;background:radial-gradient(880px 420px at 6% -14%,rgba(37,99,235,.12),transparent 62%),radial-gradient(780px 380px at 98% -8%,rgba(13,148,136,.10),transparent 60%),radial-gradient(700px 320px at 58% 112%,rgba(139,92,246,.07),transparent 65%),linear-gradient(180deg,#f9fbfe 0%,#f3f6fb 100%)}',
      '.finToolbar{display:flex;align-items:center;gap:12px;padding:13px 22px;background:rgba(255,255,255,.88);backdrop-filter:blur(16px) saturate(1.2);border-bottom:1px solid #e6ebf3;min-height:64px;box-shadow:0 1px 0 rgba(16,24,40,.02)}',
      '.finBrand{width:34px;height:34px;flex:none;border-radius:11px;display:flex;align-items:center;justify-content:center;box-shadow:0 8px 18px -10px rgba(37,99,235,.6)}',
      '.finBrand svg{display:block}',
      '.finTitle{display:flex;flex-direction:column;gap:3px;min-width:0}',
      '.finTitle b{font-size:16px;font-weight:700;letter-spacing:.2px;color:#0b1220}',
      '.finTitle span{font-size:12px;color:#64748b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:520px}',
      '.finToolbarSpacer{flex:1}',
      '.finBtn{appearance:none;border:1px solid #e2e8f2;background:#fff;color:#0f172a;border-radius:10px;padding:8px 14px;font-size:12.5px;font-weight:500;cursor:pointer;line-height:1.3;box-shadow:0 1px 1px rgba(16,24,40,.03);transition:background .15s ease,border-color .15s ease,box-shadow .15s ease}',
      '.finBtn:hover:not(:disabled){background:#f8fafc;border-color:#cbd5e1;box-shadow:0 4px 12px -10px rgba(16,24,40,.5)}',
      '.finBtn:disabled{opacity:.45;cursor:default;box-shadow:none}',
      '.finBtnPrimary{border-color:transparent;background:linear-gradient(180deg,#3b82f6,#2563eb);color:#fff;box-shadow:0 1px 2px rgba(37,99,235,.35),inset 0 1px 0 rgba(255,255,255,.18)}',
      '.finBtnPrimary:hover:not(:disabled){background:linear-gradient(180deg,#2563eb,#1d4ed8);border-color:transparent}',
      '.finBtnGhost{border-color:transparent;background:transparent;color:#64748b;box-shadow:none}',
      '.finBtnGhost:hover:not(:disabled){background:#f1f5f9;border-color:transparent}',
      '.finBody{display:flex;flex:1;min-height:0;padding:16px 16px 0;gap:16px}',
      '.finRail{width:208px;flex:none;padding:14px 11px;display:flex;flex-direction:column;gap:3px;background:rgba(255,255,255,.74);border:1px solid #e7ecf4;border-radius:18px;overflow:auto;box-shadow:0 1px 2px rgba(16,24,40,.04),0 20px 40px -36px rgba(16,24,40,.35)}',
      '.finRail::before{content:"流程";font-size:10.5px;font-weight:600;letter-spacing:1.6px;color:#94a3b8;padding:0 9px 8px}',
      '.finRailItem{display:flex;gap:10px;align-items:flex-start;padding:9px 10px;border-radius:12px;cursor:default;border:1px solid transparent;position:relative;transition:background .15s ease,border-color .15s ease}',
      '.finRailItem:hover{background:#f6f9ff}',
      '.finRailItem[data-active="true"]{background:linear-gradient(180deg,#eff6ff,#e9f2ff);border-color:#c7dbff}',
      '.finRailItem[data-active="true"]::before{content:"";position:absolute;left:-1px;top:12px;bottom:12px;width:3px;border-radius:0 3px 3px 0;background:linear-gradient(180deg,#3b82f6,#2563eb)}',
      '.finRailNum{width:22px;height:22px;flex:none;border-radius:999px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:650;background:#eef2f8;color:#64748b;box-shadow:inset 0 0 0 1px rgba(15,23,42,.04)}',
      '.finRailItem[data-active="true"] .finRailNum{background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff;box-shadow:0 5px 12px -5px rgba(37,99,235,.65)}',
      '.finRailItem[data-done="true"] .finRailNum{background:#dcfce7;color:#15803d;box-shadow:none}',
      '.finRailText{display:flex;flex-direction:column;gap:2px;min-width:0}',
      '.finRailText b{font-size:12.5px;font-weight:600;color:#0f172a}',
      '.finRailText span{font-size:11px;color:#94a3b8}',
      '.finMain{flex:1;min-width:0;overflow:auto;padding:0 4px 40px 0}',
      '.finSection{max-width:1120px;margin:0 auto 18px;background:#fff;border:1px solid #e8edf5;border-radius:18px;padding:20px 22px;box-shadow:0 1px 2px rgba(16,24,40,.04),0 24px 60px -56px rgba(16,24,40,.55)}',
      '.finSectionHead{display:flex;align-items:baseline;gap:10px;margin-bottom:14px}',
      '.finSectionHead b{font-size:15px;font-weight:700;letter-spacing:.2px;color:#0b1220}',
      '.finSectionHead b::before{content:"";display:inline-block;width:6px;height:6px;border-radius:2px;margin-right:8px;vertical-align:2px;background:linear-gradient(135deg,#3b82f6,#8b5cf6)}',
      '.finSectionHead span{font-size:12px;color:#64748b}',
      '.finEmpty{border:1px dashed #dbe3ee;border-radius:14px;padding:28px 20px;text-align:center;color:#94a3b8;line-height:1.9;background:#fbfcff}',
      '.finRow{display:flex;align-items:center;gap:10px;flex-wrap:wrap}',
      '.finInput{flex:1;min-width:220px;border:1px solid #e2e8f2;border-radius:10px;padding:9px 12px;font-size:12.5px;background:#fff;color:#0f172a;outline:none;transition:border-color .15s ease,box-shadow .15s ease}',
      '.finInput::placeholder{color:#a3adbd}',
      '.finInput:focus{border-color:#93b4ff;box-shadow:0 0 0 4px rgba(37,99,235,.10)}',
      '.finFileList{display:flex;flex-direction:column;gap:6px;margin-top:12px;max-height:240px;overflow:auto;padding-right:2px}',
      '.finFileRow{display:flex;align-items:center;gap:10px;padding:9px 12px;border:1px solid #e8edf5;border-radius:12px;cursor:pointer;background:#fff;transition:border-color .15s ease,box-shadow .15s ease,background .15s ease}',
      '.finFileRow:hover{background:#f8faff;border-color:#cfe0ff;box-shadow:0 6px 18px -14px rgba(37,99,235,.4)}',
      '.finFileRow[data-selected="true"]{border-color:#93b4ff;background:#f2f7ff;box-shadow:inset 0 0 0 1px rgba(37,99,235,.12)}',
      '.finFileRow b{font-weight:600;font-size:12.5px}',
      '.finFileRow span{font-size:11.5px;color:#8492a6}',
      '.finChip{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:4px 10px;font-size:11.5px;background:#f1f5f9;color:#526074}',
      '.finChipOk{background:#e8f8ee;color:#15803d}',
      '.finChipWarn{background:#fef4e2;color:#b45309}',
      '.finChipErr{background:#fdeaea;color:#b91c1c}',
      '.finCards{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));gap:12px;margin-bottom:16px}',
      '.finCardsDash{grid-template-columns:repeat(6,minmax(0,1fr))}',
      '@media (max-width:1280px){.finCardsDash{grid-template-columns:repeat(3,minmax(0,1fr))}}',
      '@media (max-width:840px){.finCardsDash{grid-template-columns:repeat(2,minmax(0,1fr))}}',
      '.finCard{position:relative;overflow:hidden;border:1px solid #e8edf5;border-radius:16px;padding:14px 16px;background:linear-gradient(180deg,#fff 0%,#fbfcff 100%);box-shadow:0 1px 2px rgba(16,24,40,.04),0 18px 40px -34px rgba(16,24,40,.5);transition:transform .18s ease,box-shadow .18s ease}',
      '.finCard:hover{transform:translateY(-1px);box-shadow:0 2px 6px rgba(16,24,40,.05),0 22px 44px -30px rgba(16,24,40,.45)}',
      '.finCard::before{content:"";position:absolute;left:0;right:0;top:0;height:2px;background:linear-gradient(90deg,#3b82f6,#60a5fa);opacity:.9}',
      '.finCard:nth-child(2)::before{background:linear-gradient(90deg,#f59e0b,#fbbf24)}',
      '.finCard:nth-child(3)::before{background:linear-gradient(90deg,#8b5cf6,#a78bfa)}',
      '.finCard:nth-child(4)::before{background:linear-gradient(90deg,#10b981,#34d399)}',
      '.finCard:nth-child(5)::before{background:linear-gradient(90deg,#06b6d4,#22d3ee)}',
      '.finCard:nth-child(6)::before{background:linear-gradient(90deg,#f43f5e,#fb7185)}',
      '.finCardsDash .finCard{padding:13px 14px}',
      '.finCardsDash .finCard b{font-size:20px}',
      '.finCard span{display:block;font-size:11px;color:#64748b;letter-spacing:.3px;margin-bottom:8px}',
      '.finCard b{font-size:22px;font-weight:680;letter-spacing:-.2px;font-variant-numeric:tabular-nums;color:#0b1220}',
      '.finCard em{display:block;font-style:normal;font-size:11.5px;color:#8492a6;margin-top:2px;line-height:1.2}',
      '.finTableWrap{overflow:auto;border:1px solid #e8edf5;border-radius:14px;background:#fff}',
      'table.finTable{border-collapse:collapse;width:100%;font-size:12px}',
      'table.finTable th{position:sticky;top:0;background:#f8fafc;color:#5b687b;font-weight:600;text-align:right;padding:10px 12px;border-bottom:1px solid #e8edf5;white-space:nowrap;letter-spacing:.2px}',
      'table.finTable th:first-child,table.finTable td:first-child{text-align:left}',
      'table.finTable td{padding:9px 12px;border-bottom:1px solid #f1f4f9;text-align:right;white-space:nowrap;color:#1e293b}',
      'table.finTable tr:last-child td{border-bottom:0}',
      'table.finTable tbody tr:hover td{background:#f8faff}',
      '.finUp{color:#059669;font-weight:600}',
      '.finDown{color:#dc2626;font-weight:600}',
      '.finIssue{display:flex;gap:10px;align-items:flex-start;padding:10px 0;border-bottom:1px dashed #edf1f7;line-height:1.6}',
      '.finIssue:last-child{border-bottom:0}',
      '.finIssue p{margin:0;font-size:12.5px;color:#334155}',
      '.finIssue p b{font-weight:600;color:#0f172a}',
      '.finIssue small{color:#8492a6}',
      '.finMapping{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px}',
      '.finField{display:flex;flex-direction:column;gap:7px;border:1px solid #e8edf5;border-radius:14px;padding:12px 14px;background:#fcfdff}',
      '.finField label{font-size:12.5px;font-weight:600;display:flex;align-items:center;gap:6px;color:#0f172a}',
      '.finField label em{font-style:normal;font-size:10.5px;color:#b45309;font-weight:600;background:#fef3c7;border-radius:999px;padding:1px 7px}',
      '.finField small{font-size:11.5px;color:#8492a6}',
      '.finField select{border:1px solid #e2e8f2;border-radius:9px;padding:8px 10px;font-size:12.5px;background:#fff;color:#0f172a;outline:none;transition:border-color .15s ease,box-shadow .15s ease}',
      '.finField select:focus{border-color:#93b4ff;box-shadow:0 0 0 3px rgba(37,99,235,.10)}',
      '.finDeliver{display:flex;flex-direction:column;gap:10px}',
      '.finFileCard{display:flex;align-items:center;gap:14px;border:1px solid #e8edf5;border-radius:14px;padding:12px 15px;background:linear-gradient(180deg,#fff,#fcfdff);box-shadow:0 1px 2px rgba(16,24,40,.03)}',
      '.finFileCard b{font-size:12.8px;font-weight:650;display:block;margin-bottom:3px}',
      '.finFileCard span{font-size:11.5px;color:#8492a6}',
      '.finFileIcon{width:38px;height:38px;border-radius:11px;display:flex;align-items:center;justify-content:center;font-size:10.5px;font-weight:700;flex:none;letter-spacing:.4px}',
      '.finSpark{display:flex;align-items:flex-end;gap:4px;height:78px;padding:6px 2px}',
      '.finSpark i{flex:1;min-width:6px;border-radius:5px 5px 0 0;background:linear-gradient(180deg,#93c5fd,#3b82f6);opacity:.95}',
      '.finNotice{display:flex;gap:8px;padding:11px 14px;border-radius:12px;background:#eff6ff;color:#1d4ed8;font-size:12px;line-height:1.65;margin-bottom:12px;border:1px solid #dbeafe}',
      '.finNoticeWarn{background:#fffbeb;color:#92400e;border-color:#fde68a}',
      '.finError{margin:0 auto 14px;max-width:1120px;padding:11px 15px;border-radius:12px;background:#fef2f2;color:#b91c1c;font-size:12.5px;border:1px solid #fecaca;line-height:1.6;box-shadow:0 8px 20px -18px rgba(185,28,28,.4)}',
      '.finLoading{display:inline-flex;align-items:center;gap:8px;color:#64748b;font-size:12.5px}',
      '.finSpin{width:14px;height:14px;border:2px solid #e2e8f2;border-top-color:#2563eb;border-radius:50%;animation:finspin .8s linear infinite}',
      '@keyframes finspin{to{transform:rotate(360deg)}}',
      '.finChatDrawer{width:min(420px,42%);flex:none;border-left:1px solid #e8edf5;min-height:0;display:flex;flex-direction:column;background:#fff;box-shadow:-18px 0 44px -38px rgba(16,24,40,.45)}',
      '.finChatHead{display:flex;align-items:center;gap:8px;padding:11px 14px;border-bottom:1px solid #e8edf5;font-size:12.5px;font-weight:600}',
      '.finChatBody{flex:1;min-height:0;display:flex;flex-direction:column}',
      '.finDash{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:16px}',
      '.finPanel{border:1px solid #e8edf5;border-radius:16px;padding:14px 16px 12px;background:linear-gradient(180deg,#fff,#fcfdff);min-width:0;box-shadow:0 1px 2px rgba(16,24,40,.04),0 18px 40px -36px rgba(16,24,40,.5)}',
      '.finPanelWide{grid-column:1/-1}',
      '.finPanelHead{display:flex;align-items:baseline;justify-content:space-between;gap:10px;margin-bottom:8px}',
      '.finPanelHead b{font-size:13px;font-weight:650;color:#0f172a}',
      '.finPanelHead span{font-size:11.5px;color:#94a3b8}',
      '.finChart{display:block;width:100%;height:auto}',
      '.finChart text{fill:#8492a6;font-family:inherit;font-size:11px}',
      '.finLegend{display:flex;flex-wrap:wrap;gap:4px 14px;margin-top:8px}',
      '.finLegend>span{font-size:11.5px;color:#64748b;display:inline-flex;align-items:center;gap:5px}',
      '.finLegend i,.finLegendRow i{width:9px;height:9px;border-radius:3px;display:inline-block;flex:none}',
      '.finDonut{display:flex;align-items:center;gap:16px;flex-wrap:wrap}',
      '.finLegendCol{display:flex;flex-direction:column;gap:6px;min-width:0;flex:1}',
      '.finLegendRow{display:grid;grid-template-columns:auto minmax(0,1fr) auto auto;gap:8px;align-items:center;font-size:11.8px}',
      '.finLegendName{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
      '.finLegendRow em{font-style:normal;color:#8492a6;font-variant-numeric:tabular-nums}',
      '.finLegendRow b{font-weight:600;font-variant-numeric:tabular-nums;min-width:44px;text-align:right}',
      '.finHBarRow{display:grid;grid-template-columns:minmax(60px,110px) minmax(0,1fr) auto;gap:10px;align-items:center;margin:8px 0;font-size:12px}',
      '.finHBarName{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#334155}',
      '.finHBarTrack{background:#f1f4f9;border-radius:6px;height:14px;overflow:hidden}',
      '.finHBarTrack i{display:block;height:100%;border-radius:6px}',
      '.finHBarRow b{font-weight:600;font-variant-numeric:tabular-nums;min-width:78px;text-align:right;color:#0f172a}',
      '.finPanel .finEmpty{padding:18px 12px}',
      '@media (max-width:1100px){.finDash{grid-template-columns:minmax(0,1fr)}}',
    ].join('\n');

    function useStyles() {
      useState(function () {
        if (typeof document === 'undefined') return null;
        var existing = document.getElementById('dsh-personal-finance-workbench-style');
        if (existing) return existing;
        var tag = document.createElement('style');
        tag.id = 'dsh-personal-finance-workbench-style';
        tag.textContent = CSS;
        document.head.appendChild(tag);
        return tag;
      });
    }

    function Section(props) {
      return h('section', { className: 'finSection' },
        h('div', { className: 'finSectionHead' },
          h('b', null, props.title),
          props.hint ? h('span', null, props.hint) : null,
          h('div', { style: { flex: 1 } }),
          props.extra || null),
        props.children);
    }

    function Field(props) {
      return h('div', { className: 'finField' },
        h('label', null, props.label, props.required ? h('em', null, '必填') : null),
        props.children,
        h('small', null, props.hint || ''));
    }

    function WorkbenchPanel(props) {
      useStyles();
      var ctx = host;
      var conversation = props.conversation || null;

      var [info, setInfo] = useState(null);
      var [folder, setFolder] = useState('');
      var [files, setFiles] = useState([]);
      var [filePath, setFilePath] = useState('');
      var [manualPath, setManualPath] = useState('');
      var [step, setStep] = useState('data');
      var [busy, setBusy] = useState('');
      var [error, setError] = useState('');
      var [inspectData, setInspectData] = useState(null);
      var [mapping, setMapping] = useState({});
      var [analysis, setAnalysis] = useState(null);
      var [report, setReport] = useState(null);
      var [showChat, setShowChat] = useState(false);

      useEffect(function () {
        getInfo().then(setInfo).catch(function (e) { setError(e.message); });
      }, []);

      var stepIndex = STEPS.findIndex(function (s) { return s.key === step; });

      function go(next) { setStep(next); }

      function pickDirectory() {
        if (!ctx || !ctx.uiWorkspace) {
          setError('当前还没有连接到桌面能力，暂时无法弹出文件夹选择器，可以在下方直接粘贴文件夹路径。');
          return;
        }
        setBusy('folder');
        Promise.resolve()
          .then(function () { return ctx.uiWorkspace.pickDirectory(); })
          .then(function (dir) {
            if (dir) {
              setFolder(dir);
              return post('/list', { dir: dir }).then(function (data) { setFiles(data.files || []); });
            }
            return null;
          })
          .catch(function (e) { setError(e.message); })
          .finally(function () { setBusy(''); });
      }

      function loadDirectory(dir) {
        if (!dir) return;
        setError('');
        setBusy('folder');
        post('/list', { dir: dir })
          .then(function (data) { setFolder(data.dir); setFiles(data.files || []); })
          .catch(function (e) { setError(e.message); })
          .finally(function () { setBusy(''); });
      }

      function inspectFile(target) {
        if (!target) return;
        setError('');
        setReport(null);
        setAnalysis(null);
        setFilePath(target);
        setBusy('inspect');
        post('/inspect', { path: target })
          .then(function (data) {
            setInspectData(data);
            var suggested = data.suggestedMapping || {};
            var next = {};
            MAPPING_FIELDS.forEach(function (field) {
              next[field.key] = suggested[field.key] || (field.required && data.columns && data.columns.length ? data.columns[0].name : '');
            });
            setMapping(next);
            setStep('quality');
          })
          .catch(function (e) { setError(e.message); })
          .finally(function () { setBusy(''); });
      }

      function useSample() {
        setError('');
        setBusy('sample');
        post('/sample', {})
          .then(function (data) { return inspectFile(data.path); })
          .catch(function (e) { setError(e.message); })
          .finally(function () { setBusy(''); });
      }

      function runAnalyze() {
        if (!filePath) { setError('请先选择数据文件。'); return; }
        if (!mapping.amount) { setError('请至少指定「金额列」，否则无法汇总收支。'); return; }
        setError('');
        setBusy('analyze');
        post('/analyze', { path: filePath, mapping: mapping })
          .then(function (data) { setAnalysis(data); setStep('analysis'); })
          .catch(function (e) { setError(e.message); })
          .finally(function () { setBusy(''); });
      }

      function runReport() {
        if (!analysis) { setError('请先完成分析。'); return; }
        setError('');
        setBusy('report');
        post('/report', { path: filePath, mapping: analysis.mapping || mapping })
          .then(function (data) { setReport(data); setStep('deliver'); return getInfo().then(setInfo); })
          .catch(function (e) { setError(e.message); })
          .finally(function () { setBusy(''); });
      }

      function reset() {
        setFolder('');
        setFiles([]);
        setFilePath('');
        setManualPath('');
        setInspectData(null);
        setMapping({});
        setAnalysis(null);
        setReport(null);
        setError('');
        setStep('data');
      }

      function reveal(target) {
        post('/reveal', { path: target }).catch(function (e) { setError(e.message); });
      }

      var columns = inspectData && inspectData.columns ? inspectData.columns : [];
      var kpis = analysis && analysis.kpis ? analysis.kpis : null;
      var columnsByName = useMemo(function () {
        var map = {};
        columns.forEach(function (column) { map[column.name] = column; });
        return map;
      }, [inspectData]);

      var brand = h('div', { className: 'finBrand', 'aria-hidden': 'true', style: { background: 'linear-gradient(135deg,#04211f 0%,#0f766e 62%,#14b8a6 100%)' } },
        h('svg', { width: 20, height: 20, viewBox: '0 0 20 20', fill: 'none' },
          h('circle', { cx: 10, cy: 10, r: 7, stroke: '#6ee7b7', strokeWidth: 2.4, strokeDasharray: '28 16', transform: 'rotate(-90 10 10)' }),
          h('text', { x: 10, y: 14.1, textAnchor: 'middle', fontSize: 10.5, fontWeight: 700, fill: '#ffffff', fontFamily: 'inherit' }, '¥')));

      var toolbar = h('div', { className: 'finToolbar', 'data-dsh-workbench-toolbar': '' },
        brand,
        h('div', { className: 'finTitle' },
          h('b', null, '个人财务'),
          h('span', null, filePath ? '当前数据：' + (filePath.split(/[\\/]/).pop()) : '读入个人流水 Excel / CSV，体检口径，导出可编辑的 xlsx 与一页结论 docx')),
        h('div', { className: 'finToolbarSpacer' }),
        h('button', { className: 'finBtn', onClick: pickDirectory, disabled: !!busy }, busy === 'folder' ? '正在选择…' : '选择文件夹'),
        h('button', { className: 'finBtn', onClick: useSample, disabled: !!busy }, busy === 'sample' ? '生成示例…' : '用示例数据试试'),
        filePath ? h('button', { className: 'finBtn finBtnGhost', onClick: reset }, '重新开始') : null,
        conversation ? h('button', { className: 'finBtn' + (showChat ? ' finBtnPrimary' : ''), onClick: function () { setShowChat(!showChat); } }, 'AI 对话') : null);

      var rail = h('aside', { className: 'finRail' }, STEPS.map(function (item, index) {
        var active = item.key === step;
        return h('div', {
          className: 'finRailItem',
          key: item.key,
          'data-active': active ? 'true' : 'false',
          'data-done': index < stepIndex ? 'true' : 'false',
          style: { cursor: item.key === 'data' || filePath ? 'pointer' : 'default' },
          onClick: function () { if (item.key === 'data' || filePath) go(item.key); },
        },
          h('div', { className: 'finRailNum' }, index < stepIndex ? '✓' : String(index + 1)),
          h('div', { className: 'finRailText' }, h('b', null, item.label), h('span', null, item.hint)));
      }));

      var dataSection = h(Section, { title: '数据接入', hint: '选择包含个人收支流水的 Excel 或 CSV 文件' },
        h('div', { className: 'finRow' },
          h('button', { className: 'finBtn finBtnPrimary', onClick: pickDirectory, disabled: !!busy }, '选择文件夹'),
          h('button', { className: 'finBtn', onClick: useSample, disabled: !!busy }, '使用示例数据'),
          h('span', { className: 'finChip' }, '数据只在本机读取和计算')),
        h('div', { className: 'finRow', style: { marginTop: 10 } },
          h('input', {
            className: 'finInput',
            placeholder: '也可以直接粘贴文件夹或文件路径，例如 D:\\个人账单',
            value: manualPath,
            onChange: function (event) { setManualPath(event.target.value); },
            onKeyDown: function (event) { if (event.key === 'Enter') { var value = manualPath.trim(); if (!value) return; if (/\.(xlsx|xlsm|csv|xls)$/i.test(value)) inspectFile(value); else loadDirectory(value); } },
          }),
          h('button', {
            className: 'finBtn',
            disabled: !!busy,
            onClick: function () {
              var value = manualPath.trim();
              if (!value) return;
              if (/\.(xlsx|xlsm|csv|xls)$/i.test(value)) inspectFile(value); else loadDirectory(value);
            },
          }, '读取')),
        folder ? h('div', { style: { marginTop: 12 } },
          h('div', { className: 'finRow' }, h('span', { className: 'finChip' }, shortPath(folder)),
            h('span', { className: 'finChip' }, (files.length ? files.length + ' 个数据文件' : '没有找到 xlsx / csv 文件'))),
          files.length ? h('div', { className: 'finFileList' }, files.map(function (file) {
            return h('div', {
              key: file.path,
              className: 'finFileRow',
              'data-selected': file.path === filePath ? 'true' : 'false',
              onClick: function () { inspectFile(file.path); },
            },
              h('b', null, file.name),
              h('span', { style: { flex: 1 } }, shortPath(file.path)),
              h('span', null, bytesText(file.bytes)));
          })) : h('div', { className: 'finEmpty', style: { marginTop: 10 } }, '这个文件夹里没有 xlsx / csv，换一个文件夹试试。')) : null);

      var qualitySection = null;
      if (inspectData) {
        var issues = inspectData.issues || [];
        qualitySection = h(Section, {
          title: '数据体检',
          hint: '文件：' + inspectData.fileName + (inspectData.sheet ? ' · 工作表：' + inspectData.sheet : ''),
          extra: h('button', { className: 'finBtn finBtnPrimary', onClick: function () { go('mapping'); } }, '确认口径'),
        },
          h('div', { className: 'finCards' },
            h('div', { className: 'finCard' }, h('span', null, '数据行'), h('b', null, String(inspectData.rowCount || 0))),
            h('div', { className: 'finCard' }, h('span', null, '字段数'), h('b', null, String(columns.length))),
            h('div', { className: 'finCard' }, h('span', null, '表头行'), h('b', null, String(inspectData.headerRow || 1))),
            h('div', { className: 'finCard' }, h('span', null, '提示'), h('b', null, String(issues.length)))),
          columns.length ? h('div', { className: 'finTableWrap' },
            h('table', { className: 'finTable' },
              h('thead', null, h('tr', null,
                h('th', null, '字段'), h('th', null, '类型'), h('th', null, '缺失率'), h('th', null, '示例值'))),
              h('tbody', null, columns.map(function (column) {
                var typeText = { number: '数值', date: '日期', text: '文本', empty: '空列' }[column.type] || column.type;
                var missing = column.missingRatio || 0;
                return h('tr', { key: column.name },
                  h('td', null, column.name),
                  h('td', null, typeText),
                  h('td', null, pct(missing)),
                  h('td', { style: { textAlign: 'left', color: 'var(--fin-muted,#6b7280)' } }, (column.samples || []).slice(0, 3).join(' / ')));
              })))) : null,
          issues.length ? h('div', { style: { marginTop: 12 } }, issues.map(function (issue, index) {
            var cls = issue.level === 'error' ? 'finChipErr' : issue.level === 'warn' ? 'finChipWarn' : 'finChip';
            return h('div', { className: 'finIssue', key: index },
              h('span', { className: 'finChip ' + cls }, issue.level === 'error' ? '需处理' : issue.level === 'warn' ? '注意' : '提示'),
              h('p', null, h('b', null, issue.title || ''), h('br'), h('small', null, issue.detail || '')));
          })) : h('div', { className: 'finNotice', style: { marginTop: 12 } }, '字段结构与缺失率都正常，可以进入口径确认。'));
      }

      var mappingSection = inspectData ? h(Section, {
        title: '口径映射',
        hint: '告诉工作台：哪一列是日期、金额、收支方向、分类、账户',
        extra: busy === 'analyze'
          ? h('span', { className: 'finLoading' }, h('i', { className: 'finSpin' }), '正在分析…')
          : h('button', { className: 'finBtn finBtnPrimary', onClick: runAnalyze, disabled: !!busy }, '开始分析'),
      },
        h('div', { className: 'finMapping' }, MAPPING_FIELDS.map(function (field) {
          var value = mapping[field.key] || '';
          return h(Field, { key: field.key, label: field.label, required: field.required, hint: field.hint },
            h('select', {
              value: value,
              onChange: function (event) {
                var next = Object.assign({}, mapping);
                next[field.key] = event.target.value;
                setMapping(next);
              },
            },
              h('option', { value: '' }, '（不使用）'),
              columns.map(function (column) {
                return h('option', { key: column.name, value: column.name }, column.name + ' · ' + ({ number: '数值', date: '日期', text: '文本', empty: '空' }[column.type] || column.type));
              })));
        })),
        h('div', { className: 'finNotice', style: { marginTop: 12 } }, '计算口径：净结余 = 总收入 − 总支出；储蓄率 = 净结余 ÷ 总收入；支出结构 = 分类支出 ÷ 总支出。含「合计 / 总计 / 小计」的行会自动跳过。'),
        columnsByName && mapping.amount ? h('div', { className: 'finRow', style: { marginTop: 4 } },
          mapping.date ? h('span', { className: 'finChip' }, '日期列：' + mapping.date) : null,
          h('span', { className: 'finChip' }, '金额列：' + mapping.amount),
          mapping.direction ? h('span', { className: 'finChip' }, '收支方向：' + mapping.direction) : h('span', { className: 'finChip' }, '方向：按金额正负判断'),
          mapping.category ? h('span', { className: 'finChip' }, '分类：' + mapping.category) : null,
          mapping.account ? h('span', { className: 'finChip' }, '账户：' + mapping.account) : null) : null) : null;

      var analysisSection = null;
      if (analysis && kpis) {
        var trend = analysis.trend || [];
        var expenseStructure = (analysis.expenseStructure && analysis.expenseStructure.items) || [];
        var incomeStructure = (analysis.incomeStructure && analysis.incomeStructure.items) || [];
        var anomalies = analysis.anomalies || [];
        var barsPanel = h(ChartPanel, { title: '月度收入与支出', hint: '柱高按金额 · 悬停看具体数值', wide: true },
          h(GroupedBars, {
            wide: true,
            rows: trend.map(function (item) { return { label: item.period || '', values: [item.income || 0, item.expense || 0] }; }),
            series: [{ name: '收入', color: '#22c55e' }, { name: '支出', color: '#ef4444' }],
            format: money,
          }));
        var ratePanel = h(ChartPanel, { title: '储蓄率走势', hint: '净结余 ÷ 总收入' },
          h(TrendLines, {
            zeroBase: false,
            rows: trend.map(function (item) { return { label: item.period || '', values: [item.savingsRate || 0] }; }),
            series: [{ name: '储蓄率', color: '#2563eb' }],
            format: pct,
            pctAxis: true,
          }));
        var netPanel = h(ChartPanel, { title: '月度结余', hint: '正绿负红' },
          h(GroupedBars, {
            rows: trend.map(function (item) { return { label: item.period || '', values: [item.net || 0] }; }),
            series: [{ name: '净结余', color: '#10b981' }],
            colorBySign: true,
            format: money,
          }));
        var expensePanel = h(ChartPanel, { title: '支出结构', hint: expenseStructure.length ? 'Top' + Math.min(8, expenseStructure.length) + ' 分类' : '' },
          h(DonutChart, {
            items: expenseStructure.map(function (item) { return { name: item.name, value: item.amount || 0 }; }),
            format: money,
            centerValue: expenseStructure.length ? pct(expenseStructure[0].share) : '',
            centerLabel: expenseStructure.length ? expenseStructure[0].name : '',
            maxItems: 8,
          }));
        var incomePanel = h(ChartPanel, { title: '收入来源', hint: incomeStructure.length ? 'Top' + Math.min(6, incomeStructure.length) + ' 来源' : '' },
          h(DonutChart, {
            items: incomeStructure.map(function (item) { return { name: item.name, value: item.amount || 0 }; }),
            format: money,
            centerValue: incomeStructure.length ? pct(incomeStructure[0].share) : '',
            centerLabel: incomeStructure.length ? incomeStructure[0].name : '',
            maxItems: 6,
          }));
        var anomalyPanel = h(ChartPanel, { title: '异常与提示', hint: anomalies.length + ' 条信号', wide: true },
          anomalies.length ? h('div', null, anomalies.map(function (item, index) {
            var cls = item.level === 'error' ? 'finChipErr' : item.level === 'warn' ? 'finChipWarn' : 'finChip';
            return h('div', { className: 'finIssue', key: index },
              h('span', { className: 'finChip ' + cls }, item.level === 'error' ? '异常' : item.level === 'warn' ? '注意' : '提示'),
              h('p', null, h('b', null, item.title || ''), h('br'), h('small', null, item.detail || '')));
          })) : h('div', { className: 'finEmpty' }, '没有发现明显异常，收支整体平稳。'));
        analysisSection = h(Section, {
          title: '财务看板',
          hint: (kpis.periodStart ? kpis.periodStart + ' 至 ' + kpis.periodEnd + ' · ' : '') + kpis.records + ' 条有效记录',
          extra: busy === 'report'
            ? h('span', { className: 'finLoading' }, h('i', { className: 'finSpin' }), '正在生成交付物…')
            : h('button', { className: 'finBtn finBtnPrimary', onClick: runReport, disabled: !!busy }, '生成 xlsx + docx'),
        },
          h('div', { className: 'finCards finCardsDash' },
            h('div', { className: 'finCard' }, h('span', null, '总收入'), h('b', null, money(kpis.income))),
            h('div', { className: 'finCard' }, h('span', null, '总支出'), h('b', null, money(kpis.expense))),
            h('div', { className: 'finCard' }, h('span', null, '净结余'), h('b', { className: kpis.net >= 0 ? 'finUp' : 'finDown' }, money(kpis.net)), h('em', null, '储蓄率 ' + pct(kpis.savingsRate))),
            h('div', { className: 'finCard' }, h('span', null, '储蓄率'), h('b', null, pct(kpis.savingsRate)), h('em', null, '净结余 ÷ 总收入')),
            h('div', { className: 'finCard' }, h('span', null, '月均支出'), h('b', null, money(kpis.avgMonthlyExpense)), h('em', null, kpis.months + ' 个月')),
            h('div', { className: 'finCard' }, h('span', null, '结余趋势'), h('b', { className: kpis.trendDirection >= 0 ? 'finUp' : 'finDown' }, pctDelta(kpis.trendDirection)), h('em', null, '储蓄率首末对比'))),
          h('div', { className: 'finDash' }, barsPanel, ratePanel, netPanel, expensePanel, incomePanel, anomalyPanel),
          trend.length ? h('div', { style: { marginTop: 4 } },
            h('div', { className: 'finSectionHead', style: { marginBottom: 8 } }, h('b', null, '明细数据'), h('span', null, '月度收支（xlsx 中保留完整流水）')),
            h('div', { className: 'finTableWrap' },
              h('table', { className: 'finTable' },
                h('thead', null, h('tr', null, h('th', null, '期间'), h('th', null, '收入'), h('th', null, '支出'), h('th', null, '结余'), h('th', null, '储蓄率'), h('th', null, '支出环比'))),
                h('tbody', null, trend.map(function (item) {
                  return h('tr', { key: item.period },
                    h('td', null, item.period),
                    h('td', null, money(item.income)),
                    h('td', null, money(item.expense)),
                    h('td', { className: item.net >= 0 ? 'finUp' : 'finDown' }, money(item.net)),
                    h('td', null, pct(item.savingsRate)),
                    h('td', { className: item.expenseMom == null ? '' : item.expenseMom >= 0 ? 'finDown' : 'finUp' }, item.expenseMom == null ? '—' : pctDelta(item.expenseMom)));
                }))))) : null,
          expenseStructure.length ? h('div', { style: { marginTop: 14 } },
            h('div', { className: 'finSectionHead', style: { marginBottom: 8 } }, h('b', null, '结构明细'), h('span', null, 'Top8 支出分类与笔数')),
            h('div', { className: 'finTableWrap' },
              h('table', { className: 'finTable' },
                h('thead', null, h('tr', null, h('th', null, '分类'), h('th', null, '支出金额'), h('th', null, '支出占比'), h('th', null, '笔数'))),
                h('tbody', null, expenseStructure.slice(0, 8).map(function (item) {
                  return h('tr', { key: item.name },
                    h('td', null, item.name),
                    h('td', null, money(item.amount)),
                    h('td', null, pct(item.share)),
                    h('td', null, String(item.count || 0)));
                }))))) : null);
      }

      var deliverSection = null;
      if (report) {
        var files = report.files || [];
        deliverSection = h(Section, {
          title: '交付文件',
          hint: '已生成，可下载或直接打开所在文件夹',
          extra: h('button', { className: 'finBtn', onClick: function () { reveal(info && info.exportsDir); } }, '打开输出文件夹'),
        },
          h('div', { className: 'finDeliver' }, files.map(function (file) {
            var isXlsx = file.kind === 'xlsx';
            return h('div', { className: 'finFileCard', key: file.path },
              h('div', {
                className: 'finFileIcon',
                style: isXlsx ? { background: '#dcfce7', color: '#15803d' } : { background: '#dbeafe', color: '#1d4ed8' },
              }, isXlsx ? 'XLS' : 'DOC'),
              h('div', { style: { flex: 1, minWidth: 0 } },
                h('b', null, file.name),
                h('span', null, (isXlsx ? '可编辑表格：概览、月度收支、支出结构、收入来源、收支明细、数据体检、口径卡' : '一页结论：摘要、关键指标、月度收支、支出结构、异常与口径说明') + ' · ' + bytesText(file.bytes))),
              file.url ? h('a', { className: 'finBtn finBtnPrimary', href: file.url, download: file.name, style: { textDecoration: 'none' } }, '下载') : null,
              h('button', { className: 'finBtn', onClick: function () { reveal(file.path); } }, '打开'));
          })),
          report.compare ? h('div', { className: 'finNotice', style: { marginTop: 12 } }, '已附带「复跑对比」：与这个文件上一次的分析结果对比。') : null,
          info && info.history && info.history.length ? h('div', { style: { marginTop: 14 } },
            h('div', { className: 'finSectionHead', style: { marginBottom: 8 } }, h('b', null, '最近生成'), h('span', null, '最多显示 8 条')),
            info.history.map(function (item, index) {
              return h('div', { className: 'finFileRow', key: (item.generatedAt || '') + index, style: { cursor: 'default' } },
                h('b', null, item.sourceName || '数据文件'),
                h('span', { style: { flex: 1 } }, item.generatedAt || ''),
                (item.files || []).map(function (file) {
                  return h('a', { key: file.name, className: 'finBtn', href: file.url || null, download: file.name, style: { textDecoration: 'none' } }, file.kind === 'xlsx' ? 'xlsx' : 'docx');
                }));
            })) : null);
      }

      var railHint = !filePath ? '先选择一份个人流水数据' : report ? '交付物已生成，可下载或继续复跑' : analysis ? '数据已分析，生成交付文件即可' : '数据已读入，确认口径后开始分析';

      var body = h('div', { className: 'finBody' },
        rail,
        h('div', { className: 'finMain' },
          error ? h('div', { className: 'finError' }, error) : null,
          dataSection,
          qualitySection,
          mappingSection,
          analysisSection,
          deliverSection,
          !analysis && !inspectData ? h('div', { className: 'finNotice' }, '第一步：点「选择文件夹」挑选存放个人流水的目录，或先用「示例数据」体验完整流程。数据只在本机处理，导出物保存在本机工作台目录，不上传云端。') : null),
        showChat && conversation ? h('div', { className: 'finChatDrawer' },
          h('div', { className: 'finChatHead' }, h('span', null, 'AI 对话'), h('div', { style: { flex: 1 } }), h('button', { className: 'finBtn finBtnGhost', onClick: function () { setShowChat(false); } }, '收起')),
          h('div', { className: 'finChatBody' }, conversation)) : null);

      return h('div', { className: 'finRoot' }, toolbar, body);
    }

    return {
      inject: ['desktopWorkbenches', 'uiWorkspace'],
      apply: function (ctx) {
        host = ctx;
        ctx.effect(function () {
          return ctx.desktopWorkbenches.register({
            title: '个人财务',
            repository: 'https://github.com/atlasmarkz/dsh-personal-finance-workbench',
            description: '读入 Excel / CSV 个人流水，体检口径、分析收支与储蓄率，导出可编辑 xlsx 与一页结论 docx。',
            customFrame: true,
            headerLayout: 'compact',
          }, WorkbenchPanel);
        });
      },
    };
  },
});
