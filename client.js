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

    var CSS = [
      '.finRoot{box-sizing:border-box;display:flex;flex-direction:column;width:100%;height:100%;min-height:0;background:var(--fin-page,#f6f7f9);color:var(--fin-text,#111827);font-family:var(--fin-font,-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif);font-size:13px;overflow:hidden}',
      '.finToolbar{display:flex;align-items:center;gap:14px;padding:12px 20px;background:var(--fin-panel,#ffffff);border-bottom:1px solid var(--fin-border,#e5e7eb);min-height:60px}',
      '.finTitle{display:flex;flex-direction:column;gap:2px;min-width:0}',
      '.finTitle b{font-size:15px;font-weight:650;letter-spacing:.2px}',
      '.finTitle span{font-size:12px;color:var(--fin-muted,#6b7280);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:520px}',
      '.finToolbarSpacer{flex:1}',
      '.finBtn{appearance:none;border:1px solid var(--fin-border,#d7dbe2);background:var(--fin-panel,#fff);color:var(--fin-text,#111827);border-radius:8px;padding:7px 13px;font-size:12.5px;cursor:pointer;line-height:1.3;transition:background .12s ease,border-color .12s ease}',
      '.finBtn:hover:not(:disabled){background:var(--fin-hover,#f1f3f6)}',
      '.finBtn:disabled{opacity:.5;cursor:default}',
      '.finBtnPrimary{border-color:transparent;background:#2563eb;color:#fff}',
      '.finBtnPrimary:hover:not(:disabled){background:#1d4ed8}',
      '.finBtnGhost{border-color:transparent;background:transparent;color:var(--fin-muted,#6b7280)}',
      '.finBody{display:flex;flex:1;min-height:0}',
      '.finRail{width:196px;flex:none;border-right:1px solid var(--fin-border,#e5e7eb);padding:16px 12px;display:flex;flex-direction:column;gap:4px;background:var(--fin-panel,#fff);overflow:auto}',
      '.finRailItem{display:flex;gap:10px;align-items:flex-start;padding:9px 10px;border-radius:9px;cursor:default;border:1px solid transparent}',
      '.finRailItem[data-active="true"]{background:var(--fin-active,#eef2ff);border-color:var(--fin-activeBorder,#c7d2fe)}',
      '.finRailNum{width:20px;height:20px;flex:none;border-radius:999px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;background:var(--fin-chip,#eef0f3);color:var(--fin-muted,#6b7280)}',
      '.finRailItem[data-active="true"] .finRailNum{background:#2563eb;color:#fff}',
      '.finRailItem[data-done="true"] .finRailNum{background:#dcfce7;color:#15803d}',
      '.finRailText{display:flex;flex-direction:column;gap:1px;min-width:0}',
      '.finRailText b{font-size:12.5px;font-weight:600}',
      '.finRailText span{font-size:11px;color:var(--fin-muted,#6b7280)}',
      '.finMain{flex:1;min-width:0;overflow:auto;padding:18px 22px 40px}',
      '.finSection{max-width:1040px;margin:0 auto 16px;background:var(--fin-panel,#fff);border:1px solid var(--fin-border,#e5e7eb);border-radius:12px;padding:16px 18px}',
      '.finSectionHead{display:flex;align-items:baseline;gap:10px;margin-bottom:12px}',
      '.finSectionHead b{font-size:14px;font-weight:650}',
      '.finSectionHead span{font-size:12px;color:var(--fin-muted,#6b7280)}',
      '.finEmpty{border:1px dashed var(--fin-border,#cfd5dd);border-radius:10px;padding:26px 20px;text-align:center;color:var(--fin-muted,#6b7280);line-height:1.9}',
      '.finRow{display:flex;align-items:center;gap:10px;flex-wrap:wrap}',
      '.finInput{flex:1;min-width:220px;border:1px solid var(--fin-border,#d7dbe2);border-radius:8px;padding:8px 10px;font-size:12.5px;background:var(--fin-panel,#fff);color:var(--fin-text,#111827);outline:none}',
      '.finInput:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.12)}',
      '.finFileList{display:flex;flex-direction:column;gap:6px;margin-top:12px;max-height:240px;overflow:auto}',
      '.finFileRow{display:flex;align-items:center;gap:10px;padding:8px 10px;border:1px solid var(--fin-border,#e5e7eb);border-radius:8px;cursor:pointer;background:var(--fin-panel,#fff)}',
      '.finFileRow:hover{background:var(--fin-hover,#f6f8fb)}',
      '.finFileRow[data-selected="true"]{border-color:#2563eb;background:var(--fin-active,#eef2ff)}',
      '.finFileRow b{font-weight:600;font-size:12.5px}',
      '.finFileRow span{font-size:11.5px;color:var(--fin-muted,#6b7280)}',
      '.finChip{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:3px 9px;font-size:11.5px;background:var(--fin-chip,#eef0f3);color:var(--fin-muted,#4b5563)}',
      '.finChipOk{background:#dcfce7;color:#15803d}',
      '.finChipWarn{background:#fef3c7;color:#b45309}',
      '.finChipErr{background:#fee2e2;color:#b91c1c}',
      '.finCards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}',
      '.finCard{border:1px solid var(--fin-border,#e5e7eb);border-radius:10px;padding:12px 14px;background:var(--fin-panel,#fff)}',
      '.finCard span{display:block;font-size:11.5px;color:var(--fin-muted,#6b7280);margin-bottom:6px}',
      '.finCard b{font-size:19px;font-weight:650;letter-spacing:.2px}',
      '.finCard em{font-style:normal;font-size:11.5px;color:var(--fin-muted,#6b7280);margin-left:6px}',
      '.finTableWrap{overflow:auto;border:1px solid var(--fin-border,#e5e7eb);border-radius:10px}',
      'table.finTable{border-collapse:collapse;width:100%;font-size:12px}',
      'table.finTable th{position:sticky;top:0;background:var(--fin-th,#f8fafc);color:var(--fin-muted,#4b5563);font-weight:600;text-align:right;padding:8px 10px;border-bottom:1px solid var(--fin-border,#e5e7eb);white-space:nowrap}',
      'table.finTable th:first-child,table.finTable td:first-child{text-align:left}',
      'table.finTable td{padding:7px 10px;border-bottom:1px solid var(--fin-borderSoft,#f1f3f6);text-align:right;white-space:nowrap}',
      'table.finTable tr:last-child td{border-bottom:0}',
      'table.finTable tr:hover td{background:var(--fin-hover,#f8fafc)}',
      '.finUp{color:#15803d}.finDown{color:#b91c1c}',
      '.finIssue{display:flex;gap:8px;align-items:flex-start;padding:7px 0;border-bottom:1px dashed var(--fin-border,#e5e7eb);line-height:1.55}',
      '.finIssue:last-child{border-bottom:0}',
      '.finIssue p{margin:0;font-size:12.5px}',
      '.finIssue p b{font-weight:600}',
      '.finIssue small{color:var(--fin-muted,#6b7280)}',
      '.finMapping{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px}',
      '.finField{display:flex;flex-direction:column;gap:6px;border:1px solid var(--fin-border,#e5e7eb);border-radius:10px;padding:11px 12px}',
      '.finField label{font-size:12.5px;font-weight:600;display:flex;align-items:center;gap:6px}',
      '.finField label em{font-style:normal;font-size:11px;color:#b45309;font-weight:500}',
      '.finField small{font-size:11.5px;color:var(--fin-muted,#6b7280)}',
      '.finField select{border:1px solid var(--fin-border,#d7dbe2);border-radius:7px;padding:7px 8px;font-size:12.5px;background:var(--fin-panel,#fff);color:var(--fin-text,#111827);outline:none}',
      '.finField select:focus{border-color:#2563eb}',
      '.finDeliver{display:flex;flex-direction:column;gap:8px}',
      '.finFileCard{display:flex;align-items:center;gap:12px;border:1px solid var(--fin-border,#e5e7eb);border-radius:10px;padding:11px 13px;background:var(--fin-panel,#fff)}',
      '.finFileCard b{font-size:12.5px;font-weight:600;display:block;margin-bottom:3px}',
      '.finFileCard span{font-size:11.5px;color:var(--fin-muted,#6b7280)}',
      '.finFileIcon{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex:none}',
      '.finSpark{display:flex;align-items:flex-end;gap:4px;height:78px;padding:6px 2px}',
      '.finSpark i{flex:1;min-width:6px;border-radius:3px 3px 0 0;background:linear-gradient(180deg,#60a5fa,#2563eb);opacity:.9}',
      '.finNotice{display:flex;gap:8px;padding:9px 12px;border-radius:9px;background:#eff6ff;color:#1e40af;font-size:12px;line-height:1.6;margin-bottom:12px}',
      '.finNoticeWarn{background:#fffbeb;color:#92400e}',
      '.finError{margin:0 auto 14px;max-width:1040px;padding:10px 14px;border-radius:9px;background:#fef2f2;color:#b91c1c;font-size:12.5px;border:1px solid #fecaca;line-height:1.6}',
      '.finLoading{display:inline-flex;align-items:center;gap:8px;color:var(--fin-muted,#6b7280);font-size:12.5px}',
      '.finSpin{width:14px;height:14px;border:2px solid var(--fin-border,#d7dbe2);border-top-color:#2563eb;border-radius:50%;animation:finspin .8s linear infinite}',
      '@keyframes finspin{to{transform:rotate(360deg)}}',
      '.finChatDrawer{width:min(420px,42%);flex:none;border-left:1px solid var(--fin-border,#e5e7eb);min-height:0;display:flex;flex-direction:column;background:var(--fin-panel,#fff)}',
      '.finChatHead{display:flex;align-items:center;gap:8px;padding:10px 14px;border-bottom:1px solid var(--fin-border,#e5e7eb);font-size:12.5px;font-weight:600}',
      '.finChatBody{flex:1;min-height:0;display:flex;flex-direction:column}',
      '@media (prefers-color-scheme: dark){.finRoot{--fin-page:#0f1115;--fin-panel:#171a21;--fin-border:#2a2f3a;--fin-borderSoft:#22262f;--fin-text:#e5e7eb;--fin-muted:#9aa3b2;--fin-hover:#1e232c;--fin-chip:#232833;--fin-active:#1c2536;--fin-activeBorder:#33415c;--fin-th:#1b1f27}}',
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

      var toolbar = h('div', { className: 'finToolbar', 'data-dsh-workbench-toolbar': '' },
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
        var maxExpense = Math.max.apply(null, trend.map(function (item) { return item.expense || 0; }).concat([1]));
        analysisSection = h(Section, {
          title: '分析结果',
          hint: (kpis.periodStart ? kpis.periodStart + ' 至 ' + kpis.periodEnd + ' · ' : '') + kpis.records + ' 条有效记录',
          extra: busy === 'report'
            ? h('span', { className: 'finLoading' }, h('i', { className: 'finSpin' }), '正在生成交付物…')
            : h('button', { className: 'finBtn finBtnPrimary', onClick: runReport, disabled: !!busy }, '生成 xlsx + docx'),
        },
          h('div', { className: 'finCards' },
            h('div', { className: 'finCard' }, h('span', null, '总收入'), h('b', null, money(kpis.income))),
            h('div', { className: 'finCard' }, h('span', null, '总支出'), h('b', null, money(kpis.expense))),
            h('div', { className: 'finCard' }, h('span', null, '净结余'), h('b', { className: kpis.net >= 0 ? 'finUp' : 'finDown' }, money(kpis.net)), h('em', null, '储蓄率 ' + pct(kpis.savingsRate))),
            h('div', { className: 'finCard' }, h('span', null, '储蓄率'), h('b', null, pct(kpis.savingsRate)), h('em', null, '净结余 ÷ 总收入')),
            h('div', { className: 'finCard' }, h('span', null, '月均支出'), h('b', null, money(kpis.avgMonthlyExpense)), h('em', null, kpis.months + ' 个月')),
            h('div', { className: 'finCard' }, h('span', null, '结余趋势'), h('b', { className: kpis.trendDirection >= 0 ? 'finUp' : 'finDown' }, pctDelta(kpis.trendDirection)), h('em', null, '储蓄率首末对比'))),
          trend.length ? h('div', { className: 'finSpark' }, trend.map(function (item, index) {
            var height = Math.max(6, Math.round(((item.expense || 0) / maxExpense) * 70));
            return h('i', { key: item.period || index, style: { height: height + 'px' }, title: (item.period || '') + ' · 支出 ' + money(item.expense) });
          })) : null,
          trend.length ? h('div', { style: { marginTop: 4 } },
            h('div', { className: 'finSectionHead', style: { marginBottom: 8 } }, h('b', null, '月度收支'), h('span', null, '收入、支出、结余、储蓄率与环比')),
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
            h('div', { className: 'finSectionHead', style: { marginBottom: 8 } }, h('b', null, '支出结构'), h('span', null, 'Top8 分类与占比')),
            h('div', { className: 'finTableWrap' },
              h('table', { className: 'finTable' },
                h('thead', null, h('tr', null, h('th', null, '分类'), h('th', null, '支出金额'), h('th', null, '支出占比'), h('th', null, '笔数'))),
                h('tbody', null, expenseStructure.slice(0, 8).map(function (item) {
                  return h('tr', { key: item.name },
                    h('td', null, item.name),
                    h('td', null, money(item.amount)),
                    h('td', null, pct(item.share)),
                    h('td', null, String(item.count || 0)));
                }))))) : null,
          incomeStructure.length ? h('div', { style: { marginTop: 14 } },
            h('div', { className: 'finSectionHead', style: { marginBottom: 8 } }, h('b', null, '收入来源'), h('span', null, 'Top8 来源与占比')),
            h('div', { className: 'finTableWrap' },
              h('table', { className: 'finTable' },
                h('thead', null, h('tr', null, h('th', null, '来源'), h('th', null, '收入金额'), h('th', null, '收入占比'), h('th', null, '笔数'))),
                h('tbody', null, incomeStructure.slice(0, 8).map(function (item) {
                  return h('tr', { key: item.name },
                    h('td', null, item.name),
                    h('td', null, money(item.amount)),
                    h('td', null, pct(item.share)),
                    h('td', null, String(item.count || 0)));
                }))))) : null,
          anomalies.length ? h('div', { style: { marginTop: 14 } },
            h('div', { className: 'finSectionHead', style: { marginBottom: 6 } }, h('b', null, '异常与提示'), h('span', null, anomalies.length + ' 条')),
            anomalies.map(function (item, index) {
              var cls = item.level === 'error' ? 'finChipErr' : item.level === 'warn' ? 'finChipWarn' : 'finChip';
              return h('div', { className: 'finIssue', key: index },
                h('span', { className: 'finChip ' + cls }, item.level === 'error' ? '异常' : item.level === 'warn' ? '注意' : '提示'),
                h('p', null, h('b', null, item.title || ''), h('br'), h('small', null, item.detail || '')));
            })) : null);
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
