import "./smoke.test";
import "./workbook-sheet.test";
import "./render-plan.test";
import { runStyleRegistryTest } from "./style-registry.test";
import { runBasicBlocksTest } from "./basic-blocks.test";
import { runGridBlockTest } from "./grid-block.test";
import { runTableBlockTest } from "./table-block.test";
import { runHeaderTreeTest } from "./header-tree.test";

void (async () => {
  await runStyleRegistryTest();
  await runBasicBlocksTest();
  await runGridBlockTest();
  await runTableBlockTest();
  await runHeaderTreeTest();
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
