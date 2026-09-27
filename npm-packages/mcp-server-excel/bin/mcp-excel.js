#!/usr/bin/env node

import { createLauncher } from '../lib/launcher.js';

const { main } = createLauncher({
  packageName: '@sbroenne/mcp-server-excel',
  commandName: 'excel-mcp'
});

const exitCode = main();
if (exitCode !== undefined) {
  process.exitCode = exitCode;
}
