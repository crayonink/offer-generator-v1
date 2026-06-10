// Renders the Graphviz .dot sources in this folder to SVG images under
// modules/ROOT/images/, using a pure-WASM Graphviz (no native dot, no Docker,
// no network). Run after editing any .dot:  node diagrams/render.mjs
import { Graphviz } from "@hpcc-js/wasm/graphviz";
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";

const graphviz = await Graphviz.load();
const outDir = "modules/ROOT/images";
mkdirSync(outDir, { recursive: true });

const diagrams = ["offer-generation-flow", "product-coverage-map"];
for (const name of diagrams) {
  const dot = readFileSync(`diagrams/${name}.dot`, "utf8");
  const svg = graphviz.layout(dot, "svg", "dot");
  const out = join(outDir, `${name}.svg`);
  writeFileSync(out, svg);
  console.log(`rendered ${out} (${svg.length} bytes)`);
}
