import "./smoke.test";
import "./workbook-sheet.test";
import "./render-plan.test";
import { runStyleRegistryTest } from "./style-registry.test";

void runStyleRegistryTest().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
