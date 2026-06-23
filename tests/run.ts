import "./smoke.test";
import "./workbook-sheet.test";
import "./render-plan.test";
import { runStyleRegistryTest } from "./style-registry.test";
import { runBasicBlocksTest } from "./basic-blocks.test";
import { runGridBlockTest } from "./grid-block.test";
import { runTableBlockTest } from "./table-block.test";
import { runHeaderTreeTest } from "./header-tree.test";
import { runMergeEngineTest } from "./merge-engine.test";
import { runVariableEngineTest } from "./variable-engine.test";
import { runFormulaEngineTest } from "./formula-engine.test";
import { runMultiSheetTest } from "./multi-sheet.test";

void (async () => {
  await runStyleRegistryTest();
  await runBasicBlocksTest();
  await runGridBlockTest();
  await runTableBlockTest();
  await runHeaderTreeTest();
  await runMergeEngineTest();
  await runVariableEngineTest();
  await runFormulaEngineTest();
  await runMultiSheetTest();
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
