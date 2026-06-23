import { runHeaderTreeTest } from "./header-tree.test";
import "./render-plan.test";
import "./smoke.test";
import "./workbook-sheet.test";

void (async () => {
  // await runStyleRegistryTest();
  // await runBasicBlocksTest();
  // await runGridBlockTest();
  // await runTableBlockTest();
  await runHeaderTreeTest();
  // await runMergeEngineTest();
  // await runVariableEngineTest();
  // await runFormulaEngineTest();
  // await runMultiSheetTest();
  // await runStreamingRendererTest();
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
