// validate.js — 校验生成的看板 HTML：语法 + 全部 setOption 可执行 + 逻辑链图完整。
// 用法：node validate.js <看板.html>
// 从 HTML 抽取含 `function register` 的 init 脚本，mock echarts/document/window/ResizeObserver 后运行。
const fs = require('fs');
const vm = require('vm');

const file = process.argv[2];
if (!file) { console.error('用法: node validate.js <看板.html>'); process.exit(2); }
const html = fs.readFileSync(file, 'utf8');

const blocks = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const init = blocks.find(b => b.includes('function register'));
if (!init) { console.error('[FAIL] 未找到 init 脚本'); process.exit(1); }

try {
  new vm.Script(init, { filename: 'init.js' });
  console.log('[OK] init 脚本语法检查通过');
} catch (e) {
  console.error('[FAIL] 语法错误:', e.message); process.exit(1);
}

const calls = [];
const echartsStub = {
  init: () => ({
    setOption: (opt) => {
      calls.push(opt);
      if (!opt || typeof opt !== 'object') throw new Error('setOption 收到非对象');
    },
    getDataURL: () => 'data:image/png;base64,',
    resize: () => {},
  }),
};
const fakeDom = { addEventListener() {}, querySelector: () => ({ textContent: 'x' }), closest: () => ({ querySelector: () => ({ textContent: 'x' }) }) };
const documentStub = {
  getElementById: () => fakeDom,
  querySelectorAll: () => [],
  createElement: () => ({ click() {}, set href(v) {}, set download(v) {} }),
};
const windowStub = { addEventListener() {} };
class RO { observe() {} }

const sandbox = { echarts: echartsStub, document: documentStub, window: windowStub, ResizeObserver: RO, console };
vm.createContext(sandbox);
try {
  vm.runInContext(init, sandbox, { filename: 'init.js' });
  console.log('[OK] init 执行通过，setOption 调用数:', calls.length);
} catch (e) {
  console.error('[FAIL] 运行期错误:', e.message); process.exit(1);
}

const logic = calls.find(c => c.series && c.series[0] && c.series[0].type === 'graph');
if (logic) {
  console.log('[OK] 逻辑链节点:', logic.series[0].data.length, ' 连线:', logic.series[0].links.length);
  if (logic.series[0].data.length !== 7 || logic.series[0].links.length !== 7) {
    console.error('[FAIL] 逻辑链节点/连线数异常'); process.exit(1);
  }
  const fmt = logic.tooltip.formatter;
  if (typeof fmt !== 'function') { console.error('[FAIL] tooltip.formatter 未正确注入为函数:', typeof fmt); process.exit(1); }
  console.log('[OK] tooltip.formatter 已正确注入为函数');
} else {
  console.log('[INFO] 本看板未含逻辑链图（可忽略）');
}

const types = calls.map(c => c.series && c.series[0] && c.series[0].type);
console.log('[INFO] 图表类型:', types.join(', '));
console.log('[DONE] 校验完成');
