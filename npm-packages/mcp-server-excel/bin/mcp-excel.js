#!/usr/bin/env node

import { main } from '../lib/launcher.js';

const exitCode = main();
if (exitCode !== undefined) {
  process.exitCode = exitCode;
}
