import { runBasicBlocksTest } from './basic-blocks.test';
import { runFormulaEngineTest } from './formula-engine.test';
import { runGridBlockTest } from './grid-block.test';
import { runHeaderTreeTest } from './header-tree.test';
import { runMergeEngineTest } from './merge-engine.test';
import { runMultiSheetTest } from './multi-sheet.test';
import './public-api.test';
import './render-plan.test';
import { runSectionedTableTest } from './sectioned-table.test';
import './smoke.test';
import { runStreamingRendererTest } from './streaming-renderer.test';
import { runStyleRegistryTest } from './style-registry.test';
import { runTableBlockTest } from './table-block.test';
import { runVariableEngineTest } from './variable-engine.test';
import './workbook-sheet.test';

void (async () => {
  await runStyleRegistryTest();
  await runBasicBlocksTest();
  await runGridBlockTest();
  await runTableBlockTest();
  await runHeaderTreeTest();
  await runSectionedTableTest();
  await runMergeEngineTest();
  await runVariableEngineTest();
  await runFormulaEngineTest();
  await runMultiSheetTest();
  await runStreamingRendererTest();
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
