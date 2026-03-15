/**
 * Tests for pagination implementation across all modified files.
 * Run: node test.js
 */

const fs = require('fs');
const path = require('path');

let passed = 0;
let failed = 0;

function test(name, fn) {
  try {
    fn();
    passed++;
    console.log(`  ✓ ${name}`);
  } catch (e) {
    failed++;
    console.log(`  ✗ ${name}`);
    console.log(`    ${e.message}`);
  }
}

function assert(condition, msg) {
  if (!condition) throw new Error(msg || 'Assertion failed');
}

// Load file contents
const appsScriptCode = fs.readFileSync(path.join(__dirname, 'apps-script/code.js'), 'utf8');
const rootCode = fs.readFileSync(path.join(__dirname, 'code.js'), 'utf8');
const apiJs = fs.readFileSync(path.join(__dirname, 'js/api.js'), 'utf8');
const chatJs = fs.readFileSync(path.join(__dirname, 'js/chat.js'), 'utf8');
const utilsJs = fs.readFileSync(path.join(__dirname, 'js/utils.js'), 'utf8');

// =============================================
// Phase 1: Backend — browseInventoryPaginated
// =============================================
console.log('\n--- Phase 1: Backend (apps-script/code.js) ---');

test('browseInventoryPaginated function exists', () => {
  assert(appsScriptCode.includes('function browseInventoryPaginated(params)'),
    'Missing browseInventoryPaginated function');
});

test('doPost routes browseInventoryPaginated', () => {
  assert(appsScriptCode.includes("case 'browseInventoryPaginated':"),
    'Missing doPost case for browseInventoryPaginated');
});

test('accepts page parameter', () => {
  assert(appsScriptCode.includes('params.page'),
    'Function should accept page param');
});

test('accepts pageSize parameter', () => {
  assert(appsScriptCode.includes('params.pageSize'),
    'Function should accept pageSize param');
});

test('accepts query parameter for filtering', () => {
  assert(appsScriptCode.includes('params.query'),
    'Function should accept query param');
});

test('accepts sortColumn parameter', () => {
  assert(appsScriptCode.includes('params.sortColumn'),
    'Function should accept sortColumn param');
});

test('accepts sortDirection parameter', () => {
  assert(appsScriptCode.includes('params.sortDirection'),
    'Function should accept sortDirection param');
});

test('uses CacheService for caching', () => {
  const fn = appsScriptCode.substring(
    appsScriptCode.indexOf('function browseInventoryPaginated'),
    appsScriptCode.indexOf('function browseInventoryPaginated') + 3000
  );
  assert(fn.includes('CacheService.getScriptCache()'),
    'Should use CacheService for caching');
});

test('returns paginated response shape', () => {
  const fnStart = appsScriptCode.indexOf('function browseInventoryPaginated');
  const fnEnd = appsScriptCode.indexOf('\n// ====', fnStart + 1);
  const fn = appsScriptCode.substring(fnStart, fnEnd);
  assert(fn.includes('totalPages'), 'Response should include totalPages');
  assert(fn.includes('pageSize'), 'Response should include pageSize');
});

test('server-side filtering with query', () => {
  const fn = appsScriptCode.substring(
    appsScriptCode.indexOf('function browseInventoryPaginated'),
    appsScriptCode.indexOf('function browseInventoryPaginated') + 3000
  );
  assert(fn.includes('filtered') && fn.includes('.filter('),
    'Should filter items server-side');
});

test('server-side sorting', () => {
  const fn = appsScriptCode.substring(
    appsScriptCode.indexOf('function browseInventoryPaginated'),
    appsScriptCode.indexOf('function browseInventoryPaginated') + 3000
  );
  assert(fn.includes('filtered.sort('),
    'Should sort items server-side');
});

test('paginates with slice', () => {
  const fnStart = appsScriptCode.indexOf('function browseInventoryPaginated');
  const fnEnd = appsScriptCode.indexOf('\n// ====', fnStart + 1);
  const fn = appsScriptCode.substring(fnStart, fnEnd);
  assert(fn.includes('.slice(start, start + pageSize)'),
    'Should paginate using slice');
});

test('original browseInventory still exists', () => {
  assert(appsScriptCode.includes('function browseInventory()'),
    'Original browseInventory should be preserved');
});

test('original browseInventory doPost route still exists', () => {
  assert(appsScriptCode.includes("case 'browseInventory':"),
    'Original doPost route should be preserved');
});

test('stock status calculation preserved in paginated version', () => {
  const fn = appsScriptCode.substring(
    appsScriptCode.indexOf('function browseInventoryPaginated'),
    appsScriptCode.indexOf('function browseInventoryPaginated') + 3000
  );
  assert(fn.includes('isLowStock') && fn.includes('isCritical'),
    'Should calculate stock status in paginated version');
});

// =============================================
// Phase 1b: Root code.js mirrors apps-script
// =============================================
console.log('\n--- Phase 1b: Root code.js (mirror) ---');

test('root code.js has browseInventoryPaginated function', () => {
  assert(rootCode.includes('function browseInventoryPaginated(params)'),
    'Missing browseInventoryPaginated in root code.js');
});

test('root code.js has doPost route', () => {
  assert(rootCode.includes("case 'browseInventoryPaginated':"),
    'Missing doPost case in root code.js');
});

test('root code.js uses CacheService', () => {
  const fn = rootCode.substring(
    rootCode.indexOf('function browseInventoryPaginated'),
    rootCode.indexOf('function browseInventoryPaginated') + 3000
  );
  assert(fn.includes('CacheService.getScriptCache()'),
    'Root code.js should use CacheService');
});

// =============================================
// Phase 4: Frontend API method
// =============================================
console.log('\n--- Phase 4: Frontend API (js/api.js) ---');

test('browseInventoryPaginated method exists in APIManager', () => {
  assert(apiJs.includes('async browseInventoryPaginated(params)'),
    'Missing browseInventoryPaginated method in APIManager');
});

test('calls correct Google Script function', () => {
  assert(apiJs.includes("'browseInventoryPaginated'"),
    'Should call browseInventoryPaginated on backend');
});

test('passes params to backend', () => {
  assert(apiJs.includes('[params]'),
    'Should pass params array to backend');
});

test('original browseInventory method preserved', () => {
  assert(apiJs.includes('async browseInventory()'),
    'Original browseInventory should be preserved');
});

test('error handling returns correct shape', () => {
  const fnStart = apiJs.indexOf('async browseInventoryPaginated');
  const fnBlock = apiJs.substring(fnStart, fnStart + 500);
  assert(fnBlock.includes('totalPages'),
    'Error response should include totalPages');
});

// =============================================
// Phase 2: Frontend pagination UI & state
// =============================================
console.log('\n--- Phase 2: Frontend Pagination (js/chat.js) ---');

test('handleBrowseInventory initializes pagination state', () => {
  assert(chatJs.includes('this._browsePage = 1'), 'Should init _browsePage');
  assert(chatJs.includes('this._browsePageSize = 50'), 'Should init _browsePageSize');
  assert(chatJs.includes('this._browseTotalPages = 1'), 'Should init _browseTotalPages');
  assert(chatJs.includes("this._browseQuery = ''"), 'Should init _browseQuery');
});

test('handleBrowseInventory calls paginated API', () => {
  assert(chatJs.includes('api.browseInventoryPaginated'),
    'Should call browseInventoryPaginated');
});

test('buildInventoryTable renders pagination controls', () => {
  assert(chatJs.includes('inventory-pagination'),
    'Should render pagination container');
  assert(chatJs.includes('inventory-page-btn'),
    'Should render pagination buttons');
});

test('buildInventoryTable renders Prev button', () => {
  assert(chatJs.includes('Prev</button>'),
    'Should render Prev button');
});

test('buildInventoryTable renders Next button', () => {
  assert(chatJs.includes('Next</button>'),
    'Should render Next button');
});

test('buildInventoryTable shows page indicator', () => {
  assert(chatJs.includes('Page ${this._browsePage} of ${this._browseTotalPages}'),
    'Should show page X of Y');
});

test('goToInventoryPage method exists', () => {
  assert(chatJs.includes('goToInventoryPage(page)'),
    'Should have goToInventoryPage method');
});

test('goToInventoryPage guards against out-of-range', () => {
  assert(chatJs.includes('page < 1') && chatJs.includes('page > this._browseTotalPages'),
    'Should guard against invalid page numbers');
});

test('goToInventoryPage guards against double-click', () => {
  assert(chatJs.includes('if (this._browseLoading) return'),
    'goToInventoryPage should check _browseLoading');
});

test('Prev button disabled on page 1', () => {
  assert(chatJs.includes("this._browsePage <= 1 ? 'disabled'"),
    'Prev should be disabled on first page');
});

test('Next button disabled on last page', () => {
  assert(chatJs.includes("this._browsePage >= this._browseTotalPages ? 'disabled'"),
    'Next should be disabled on last page');
});

// =============================================
// Phase 2b: Sort now triggers server-side
// =============================================
console.log('\n--- Phase 2b: Server-side sorting (js/chat.js) ---');

test('sortInventoryTable resets to page 1', () => {
  const sortFn = chatJs.substring(
    chatJs.indexOf('sortInventoryTable(column)'),
    chatJs.indexOf('sortInventoryTable(column)') + 500
  );
  assert(sortFn.includes('this._browsePage = 1'),
    'Sort should reset to page 1');
});

test('sortInventoryTable calls _fetchInventoryPage', () => {
  const sortFn = chatJs.substring(
    chatJs.indexOf('sortInventoryTable(column)'),
    chatJs.indexOf('sortInventoryTable(column)') + 500
  );
  assert(sortFn.includes('this._fetchInventoryPage()'),
    'Sort should trigger server-side fetch');
});

test('sortInventoryTable no longer sorts client-side', () => {
  const sortFn = chatJs.substring(
    chatJs.indexOf('sortInventoryTable(column)'),
    chatJs.indexOf('sortInventoryTable(column)') + 500
  );
  assert(!sortFn.includes('this._browseItems.sort('),
    'Sort should NOT do client-side array sort');
});

test('sort direction toggles correctly', () => {
  const sortFn = chatJs.substring(
    chatJs.indexOf('sortInventoryTable(column)'),
    chatJs.indexOf('sortInventoryTable(column)') + 500
  );
  assert(sortFn.includes('this._browseSortAsc = !this._browseSortAsc'),
    'Should toggle sort direction on same column');
});

// =============================================
// Phase 3: Debounced server-side filtering
// =============================================
console.log('\n--- Phase 3: Debounced Filtering (js/chat.js) ---');

test('debounce utility exists in utils.js', () => {
  assert(utilsJs.includes('debounce(func, wait)'),
    'PerformanceUtils.debounce should exist');
});

test('creates debounced filter function', () => {
  assert(chatJs.includes('_debouncedBrowseFilter = PerformanceUtils.debounce'),
    'Should create debounced filter using PerformanceUtils.debounce');
});

test('debounce wait is 300ms', () => {
  assert(chatJs.includes('}, 300)'),
    'Debounce should use 300ms wait');
});

test('filter resets to page 1', () => {
  const match = chatJs.match(/_debouncedBrowseFilter\s*=[\s\S]*?this\._browsePage = 1/);
  assert(match, 'Filter should reset to page 1');
});

test('filter calls _fetchInventoryPage', () => {
  const debounceBlock = chatJs.substring(
    chatJs.indexOf('_debouncedBrowseFilter = PerformanceUtils.debounce'),
    chatJs.indexOf('_debouncedBrowseFilter = PerformanceUtils.debounce') + 300
  );
  assert(debounceBlock.includes('this._fetchInventoryPage()'),
    'Filter should call _fetchInventoryPage');
});

test('oninput calls debounced filter', () => {
  assert(chatJs.includes('oninput="window.app.chat._debouncedBrowseFilter(this.value)"'),
    'Input should call _debouncedBrowseFilter');
});

test('legacy filterInventoryTable still exists', () => {
  assert(chatJs.includes('filterInventoryTable(query)'),
    'filterInventoryTable should still exist as fallback');
});

// =============================================
// Phase 5: Loading states
// =============================================
console.log('\n--- Phase 5: Loading States (js/chat.js) ---');

test('_fetchInventoryPage has loading guard', () => {
  const fn = chatJs.substring(
    chatJs.indexOf('async _fetchInventoryPage()'),
    chatJs.indexOf('async _fetchInventoryPage()') + 2000
  );
  assert(fn.includes('if (this._browseLoading) return'),
    'Should guard against concurrent requests');
});

test('sets _browseLoading true at start', () => {
  const fn = chatJs.substring(
    chatJs.indexOf('async _fetchInventoryPage()'),
    chatJs.indexOf('async _fetchInventoryPage()') + 300
  );
  assert(fn.includes('this._browseLoading = true'),
    'Should set loading true at start');
});

test('resets _browseLoading in finally block', () => {
  const fnStart = chatJs.indexOf('async _fetchInventoryPage()');
  const fnEnd = chatJs.indexOf('\n    buildInventoryTable', fnStart);
  const fn = chatJs.substring(fnStart, fnEnd);
  assert(fn.includes('finally') && fn.includes('this._browseLoading = false'),
    'Should reset loading in finally block');
});

test('shows loading overlay', () => {
  const fn = chatJs.substring(
    chatJs.indexOf('async _fetchInventoryPage()'),
    chatJs.indexOf('async _fetchInventoryPage()') + 2000
  );
  assert(fn.includes('inventory-loading-overlay'),
    'Should show loading overlay');
});

test('disables pagination buttons during load', () => {
  const fn = chatJs.substring(
    chatJs.indexOf('async _fetchInventoryPage()'),
    chatJs.indexOf('async _fetchInventoryPage()') + 2000
  );
  assert(fn.includes("btn.disabled = true"),
    'Should disable pagination buttons during load');
});

test('esc() HTML escaping preserved', () => {
  assert(chatJs.includes("const esc = (text) => {"),
    'esc() function should be preserved in buildInventoryTable');
});

// =============================================
// Cross-cutting concerns
// =============================================
console.log('\n--- Cross-cutting ---');

test('_fetchInventoryPage sends sortColumn to backend', () => {
  const fn = chatJs.substring(
    chatJs.indexOf('async _fetchInventoryPage()'),
    chatJs.indexOf('async _fetchInventoryPage()') + 2000
  );
  assert(fn.includes('sortColumn: this._browseSortColumn'),
    'Should send sortColumn to backend');
});

test('_fetchInventoryPage sends sortDirection to backend', () => {
  const fn = chatJs.substring(
    chatJs.indexOf('async _fetchInventoryPage()'),
    chatJs.indexOf('async _fetchInventoryPage()') + 2000
  );
  assert(fn.includes("this._browseSortAsc ? 'asc' : 'desc'"),
    'Should send sortDirection to backend');
});

test('_fetchInventoryPage sends query to backend', () => {
  const fn = chatJs.substring(
    chatJs.indexOf('async _fetchInventoryPage()'),
    chatJs.indexOf('async _fetchInventoryPage()') + 2000
  );
  assert(fn.includes('query: this._browseQuery'),
    'Should send query to backend');
});

test('filter input preserves value after page refresh', () => {
  assert(chatJs.includes('filterInput.value = this._browseQuery'),
    'Should restore filter input value after re-render');
});

test('sort indicators rendered inline in buildInventoryTable', () => {
  const buildStart = chatJs.indexOf('buildInventoryTable(items, total)');
  const buildEnd = chatJs.indexOf('\n    goToInventoryPage', buildStart);
  const buildFn = chatJs.substring(buildStart, buildEnd);
  assert(buildFn.includes("this._browseSortColumn === col.key"),
    'Should render sort indicators inline');
});

// =============================================
// Results
// =============================================
console.log(`\n${'='.repeat(45)}`);
console.log(`Results: ${passed} passed, ${failed} failed, ${passed + failed} total`);
console.log(`${'='.repeat(45)}\n`);

process.exit(failed > 0 ? 1 : 0);
