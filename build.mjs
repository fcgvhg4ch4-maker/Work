// Inlines data/products.json into src/page.html -> index.html
import { readFileSync, writeFileSync } from "node:fs";

const data = readFileSync("data/products.json", "utf8");
JSON.parse(data); // fail fast on malformed data

const tpl = readFileSync("src/page.html", "utf8");
if (!tpl.includes("__DATA__")) throw new Error("template is missing the __DATA__ placeholder");

// </script> inside a JSON payload would close the host <script> tag early.
const safe = JSON.stringify(JSON.parse(data)).replace(/<\/script/gi, "<\\/script");

writeFileSync("index.html", tpl.replace("__DATA__", safe));
console.log(`built index.html (${(tpl.length + safe.length) / 1024 | 0} KB)`);
