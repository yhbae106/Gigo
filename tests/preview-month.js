const assert = require('node:assert/strict');
const vm = require('node:vm');

module.exports = async function verifyPreviewMonth(script) {
  // The cumulative worksheet has distinct amounts in each month's real columns.
  const values = {
    6: {1: '도매상', 6: '10월 수금대상액', 7: '최종 수금액', 9: '11월 수금대상액', 10: '최종 수금액', 15: '1월 수금대상액', 16: '최종 수금액'},
    7: {1: 10093445, 6: 1234, 7: 2000, 9: 5678, 10: 6000, 15: -3456, 16: 0},
    8: {9: 7890},
    9: {1: '숨김 행'}
  };
  const sheet = {
    getCell: (r, c) => ({value: values[r]?.[c]}),
    getColumn: () => ({width: 15}),
    getRow: r => ({hidden: r === 9})
  };
  const drawn = [];
  const canvasContext = {
    fillRect() {}, strokeRect() {}, setTransform() {},
    measureText: text => ({width: text.length}),
    fillText: text => drawn.push(text)
  };
  const elements = new Map();
  function getElementById(id) {
    if (!elements.has(id)) elements.set(id, {style: {}, addEventListener() {}});
    return elements.get(id);
  }
  const canvas = getElementById('previewCanvas');
  canvas.getContext = () => canvasContext;
  canvas.toDataURL = () => '';
  const app = vm.createContext({document: {getElementById}, window: {}, console, sessionStorage: {getItem: () => null}});
  vm.runInContext(script, app);
  const base = {wb: {getWorksheet: () => sheet}};
  app.loadBaseWorkbook = async () => base;
  // Values are already populated. Use the same column mapping as the writer.
  app.applyMonth = (ws, result, month) => app.monthCols(month);

  const cases = [
    ['2026-10', '10월 수금대상액', ['1,234', '2,000']],
    ['2026-11', '11월 수금대상액', ['5,678', '6,000', '7,890']],
    ['2027-01', '1월 수금대상액', ['-3,456', '-']]
  ];
  for (const [month, header, amounts] of cases) {
    drawn.length = 0;
    await app.renderExcelPreview({}, month);
    assert.deepEqual(drawn, ['도매상', header, '최종 수금액', '10093445', ...amounts], month);
  }

  // A delayed October restore must not overwrite a newer November upload.
  let finishRestore;
  const delayed = new Promise(resolve => { finishRestore = resolve; });
  app.loadBaseWorkbook = month => month === '2026-10' ? delayed : base;
  drawn.length = 0;
  const october = app.renderExcelPreview({}, '2026-10');
  await app.renderExcelPreview({}, '2026-11');
  finishRestore(base);
  await october;
  assert.deepEqual(drawn.filter(text => text.includes('월 수금대상액')), ['11월 수금대상액']);
};
