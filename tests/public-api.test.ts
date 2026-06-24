import assert from 'node:assert/strict';
import * as publicApi from '../src';
import * as advancedApi from '../src/advanced';

assert.equal(Object.prototype.hasOwnProperty.call(publicApi, 'RenderPlanBuilder'), false);
assert.equal(Object.prototype.hasOwnProperty.call(publicApi, 'normalizeMergeRange'), false);
assert.equal(Object.prototype.hasOwnProperty.call(publicApi, 'collectFormulaDependencies'), false);
assert.equal(Object.prototype.hasOwnProperty.call(publicApi, 'createFormulaId'), false);

assert.equal(typeof publicApi.defineWorkbook, 'function');
assert.equal(typeof publicApi.renderWorkbook, 'function');
assert.equal(typeof publicApi.compileWorkbookToRenderPlan, 'function');

assert.equal(typeof advancedApi.RenderPlanBuilder, 'function');
assert.equal(typeof advancedApi.normalizeMergeRange, 'function');
assert.equal(typeof advancedApi.collectFormulaDependencies, 'function');
assert.equal(typeof advancedApi.createFormulaId, 'function');
