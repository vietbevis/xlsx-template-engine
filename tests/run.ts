import "./smoke.test";
import "./workbook-sheet.test";
import "./render-plan.test";
import { runStyleRegistryTest } from "./style-registry.test";
import { runBasicBlocksTest } from "./basic-blocks.test";

void (async () => {
  await runStyleRegistryTest();
  await runBasicBlocksTest();
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
