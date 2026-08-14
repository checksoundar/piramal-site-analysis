/* Group urls-all.json into urls-grouped.json per the excat-url-discovery reference algorithm. */
const fs = require("fs");
const CATALOG = process.argv[2];
const siteUrl = process.argv[3];
const data = JSON.parse(fs.readFileSync(CATALOG + "/urls-all.json", "utf8"));
let urls = data["analysis-urls-all"].urls.map((u) => u.url);
if (!urls.includes(siteUrl) && !urls.includes(siteUrl + "/")) urls.push(siteUrl + "/");

const LOCALES = ["ja", "en", "id", "zh", "ko", "th", "vi", "pt", "es", "ar"];

function extractPattern(url) {
  const p = new URL(url).pathname.toLowerCase();
  if (p === "/" || p === "") return "/";
  const segs = p.split("/").filter((s) => s);
  if (segs.length === 1) return "/" + segs[0];
  let dir = segs.slice(0, -1);
  while (dir.length > 0 && /^\d+$/.test(dir[dir.length - 1])) dir = dir.slice(0, -1);
  if (dir.length === 0) return "/" + segs[0];
  return "/" + dir.join("/");
}

urls.sort(
  (a, b) =>
    new URL(b).pathname.split("/").filter((s) => s).length -
    new URL(a).pathname.split("/").filter((s) => s).length,
);

const map = {};
for (const u of urls) {
  const pat = extractPattern(u);
  (map[pat] = map[pat] || []).push(u);
}

const locales = {};
for (const u of urls) {
  const segs = new URL(u).pathname.split("/").filter((s) => s).slice(0, -1);
  let loc = "unknown";
  for (const s of segs) {
    if (LOCALES.includes(s)) { loc = s; break; }
  }
  locales[loc] = (locales[loc] || 0) + 1;
}

const groups = {};
const singles = [];
for (const [pat, list] of Object.entries(map)) {
  if (pat === "/" || list.length >= 2) {
    groups[pat] = {
      confidence: list.length >= 5 ? "95%" : list.length >= 2 ? "85%" : "100%",
      urls: list.map((url) => ({ url })),
    };
  } else singles.push(...list);
}
const unknown = [];
for (const u of singles) {
  const segs = new URL(u).pathname.split("/").filter((s) => s);
  let placed = false;
  for (let i = segs.length - 1; i > 0; i--) {
    const parent = "/" + segs.slice(0, i).join("/");
    if (groups[parent]) { groups[parent].urls.push({ url: u }); placed = true; break; }
  }
  if (!placed) unknown.push({ url: u });
}
if (unknown.length) groups["unknown"] = { confidence: "50%", urls: unknown };

const out = {
  "analysis-urls-grouped": {
    captured: new Date().toISOString(),
    urlGroupings: Object.keys(groups).length,
    locales,
    groups,
  },
};
fs.writeFileSync(CATALOG + "/urls-grouped.json", JSON.stringify(out, null, 2));
const total = Object.values(groups).reduce((a, g) => a + g.urls.length, 0);
console.log(`${process.argv[4]}: ${Object.keys(groups).length} groups, ${total} urls, locales ${JSON.stringify(locales)}`);
