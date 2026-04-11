#!/usr/bin/env node

import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");

const args = process.argv.slice(2);
const opts = parseArgs(args);

const scripts = [
  {
    name: "Check people connector status",
    file: "tools/debug/check-people-connector-status.mjs",
    needsConnectionId: true,
  },
  { name: "Verify ingestion progress", file: "tools/debug/verify-ingestion-progress.mjs" },
  {
    name: "Check app permissions",
    file: "tools/debug/check-app-permissions.mjs",
    needsConnectionId: true,
  },
];

const env = buildEnv(opts);

const results = [];

await runChecklist();

async function runChecklist() {
  console.log("Running debug checklist...");
  if (!opts.tenantId) {
    console.warn("WARN: --tenant-id not provided. Some checks may fail.");
  }
  if (!opts.connectionId) {
    console.warn("WARN: --connection-id not provided. Some checks may fail.");
  }

  for (const script of scripts) {
    const scriptPath = path.resolve(repoRoot, script.file);
    if (!fs.existsSync(scriptPath)) {
      results.push({ name: script.name, status: "SKIP", code: null });
      console.warn(`SKIP: ${script.name} (missing ${script.file})`);
      continue;
    }

    const argv = buildArgv(script);
    const code = await runNodeScript(scriptPath, argv, env);
    const status = code === 0 ? "PASS" : "FAIL";
    results.push({ name: script.name, status, code });
    if (status === "FAIL" && opts.strict) {
      break;
    }
  }

  printSummary(results);

  const failed = results.some((result) => result.status === "FAIL");
  if (failed) {
    process.exitCode = 1;
  }
}

function parseArgs(argv) {
  const options = {
    tenantId: null,
    connectionId: null,
    strict: false,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--tenant-id") {
      options.tenantId = argv[i + 1] || null;
      i += 1;
      continue;
    }
    if (arg === "--connection-id") {
      options.connectionId = argv[i + 1] || null;
      i += 1;
      continue;
    }
    if (arg === "--strict") {
      options.strict = true;
    }
  }

  return options;
}

function buildEnv(options) {
  const env = { ...process.env };

  if (options.tenantId) {
    env.TENANT_ID = options.tenantId;
    env.M365_TENANT_ID = options.tenantId;
    env.AZURE_TENANT_ID = options.tenantId;
  }

  if (options.connectionId) {
    env.CONNECTION_ID = options.connectionId;
    env.M365_CONNECTION_ID = options.connectionId;
  }

  return env;
}

function runNodeScript(scriptPath, argv, env) {
  return new Promise((resolve) => {
    console.log(`\n==> ${path.relative(repoRoot, scriptPath)}`);
    const child = spawn(process.execPath, [scriptPath, ...argv], {
      cwd: repoRoot,
      env,
      stdio: "inherit",
    });

    child.on("close", (code) => resolve(code ?? 1));
  });
}

function buildArgv(script) {
  const argv = [];
  if (script.needsConnectionId && opts.connectionId) {
    argv.push(opts.connectionId);
  }
  return argv;
}

function printSummary(items) {
  console.log("\nChecklist summary:");
  for (const item of items) {
    const code = item.code === null ? "-" : String(item.code);
    console.log(`- ${item.status} | ${item.name} | exit ${code}`);
  }
}
