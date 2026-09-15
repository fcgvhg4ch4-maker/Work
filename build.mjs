// Inlines data/products.json into src/page.html -> index.html
import { readFileSync, writeFileSync } from "node:fs";

// The Excel book ships inside the page: the viewer sandbox blocks page-initiated
// downloads, so the bytes have to be on hand for the downloads capability.
const XLSX_PATH = "dist/日本の生保_貯蓄性商品ラインナップ横断表.xlsx";

const data = readFileSync("data/products.json", "utf8");
JSON.parse(data); // fail fast on malformed data

const tpl = readFileSync("src/page.html", "utf8");
if (!tpl.includes("__DATA__")) throw new Error("template is missing the __DATA__ placeholder");

// </script> inside a JSON payload would close the host <script> tag early.
const safe = JSON.stringify(JSON.parse(data)).replace(/<\/script/gi, "<\\/script");

if (!tpl.includes("__XLSX_B64__")) throw new Error("template is missing the __XLSX_B64__ placeholder");
const xlsxB64 = readFileSync(XLSX_PATH).toString("base64");

const html = tpl.replace("__DATA__", safe).replace("__XLSX_B64__", xlsxB64);
writeFileSync("index.html", html);
console.log(`built index.html (${html.length / 1024 | 0} KB, xlsx ${xlsxB64.length / 1024 | 0} KB base64)`);
