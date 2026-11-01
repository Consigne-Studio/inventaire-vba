#!/usr/bin/env node
import { principal } from '../src/cli.mjs';

principal(process.argv.slice(2)).then((code) => { process.exitCode = code; }, (e) => {
  process.stderr.write(`Erreur : ${e.message}\n`);
  process.exitCode = 1;
});
