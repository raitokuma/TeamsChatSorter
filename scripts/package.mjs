// 產生 Teams 應用程式套件：build/TeamsChatSorter.zip（manifest.json + 兩個圖示）
// 用法：npm run package   （設定讀自 .env 或環境變數）
import { readFileSync, writeFileSync, mkdirSync, existsSync, appendFileSync } from "node:fs";
import { deflateRawSync } from "node:zlib";
import { randomUUID } from "node:crypto";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const envPath = resolve(root, ".env");

const env = {};
if (existsSync(envPath)) {
  for (const line of readFileSync(envPath, "utf8").split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (m) env[m[1]] = m[2].replace(/^["']|["']$/g, "");
  }
}
const get = (k, d) => process.env[k] || env[k] || d;

let teamsAppId = get("TEAMS_APP_ID");
if (!teamsAppId) {
  teamsAppId = randomUUID();
  appendFileSync(envPath, `\nTEAMS_APP_ID=${teamsAppId}\n`);
  console.log(`已產生 TEAMS_APP_ID=${teamsAppId} 並寫入 .env（之後請保持不變）`);
}

let host = get("APP_HOST_URL");
const clientId = get("VITE_CLIENT_ID");
if (!host || !/^https:\/\//.test(host)) throw new Error("請在 .env 設定 APP_HOST_URL（https:// 開頭，例如 https://you.github.io/TeamsChatSorter/）");
if (!clientId || clientId.startsWith("00000000")) throw new Error("請在 .env 設定 VITE_CLIENT_ID");
if (!host.endsWith("/")) host += "/";

const vars = {
  TEAMS_APP_ID: teamsAppId,
  APP_VERSION: get("APP_VERSION", "1.0.0"),
  DEVELOPER_NAME: get("DEVELOPER_NAME", "Internal"),
  APP_HOST_URL: host,
  APP_HOST_DOMAIN: new URL(host).host,
  VITE_CLIENT_ID: clientId,
};

const template = readFileSync(resolve(root, "appPackage/manifest.template.json"), "utf8");
const manifest = template.replace(/\{\{(\w+)\}\}/g, (_, k) => {
  if (!(k in vars)) throw new Error(`manifest 缺少變數 ${k}`);
  return vars[k];
});
JSON.parse(manifest); // 驗證 JSON

// ---- 最小 ZIP 寫入器（不需任何套件）----
const crcTable = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
function zip(files) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const { name, data } of files) {
    const nameBuf = Buffer.from(name, "utf8");
    const comp = deflateRawSync(data);
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6);
    lh.writeUInt16LE(8, 8); lh.writeUInt32LE(0, 10); lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26); lh.writeUInt16LE(0, 28);
    locals.push(lh, nameBuf, comp);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(8, 10); ch.writeUInt32LE(0, 12);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28); ch.writeUInt32LE(offset, 42);
    centrals.push(ch, nameBuf);
    offset += lh.length + nameBuf.length + comp.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

const out = resolve(root, "build");
mkdirSync(out, { recursive: true });
writeFileSync(resolve(out, "manifest.json"), manifest);
const pkg = zip([
  { name: "manifest.json", data: Buffer.from(manifest, "utf8") },
  { name: "color.png", data: readFileSync(resolve(root, "appPackage/color.png")) },
  { name: "outline.png", data: readFileSync(resolve(root, "appPackage/outline.png")) },
]);
writeFileSync(resolve(out, "TeamsChatSorter.zip"), pkg);
console.log(`✔ 已產生 build/TeamsChatSorter.zip（App 網址：${host}）`);
