import { readFile, rename, writeFile } from "node:fs/promises";
import path from "node:path";

const commit = process.env.VERCEL_GIT_COMMIT_SHA;

if (!commit) {
  console.log("Deployment status stamp skipped outside a Vercel Git build.");
  process.exit(0);
}

const statusPath = path.join(process.cwd(), "public", "data", "update-status.json");
const temporaryPath = `${statusPath}.tmp`;
const status = JSON.parse(await readFile(statusPath, "utf8"));

status.sourceCommit = commit;
status.deploymentStatus = "deployed";
status.lastProductionDeploymentAt = new Date().toISOString();

await writeFile(temporaryPath, `${JSON.stringify(status, null, 2)}\n`, "utf8");
await rename(temporaryPath, statusPath);
console.log(`Stamped production deployment status for ${commit}.`);
