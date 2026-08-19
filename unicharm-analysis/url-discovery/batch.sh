#!/usr/bin/env bash
# Batch URL discovery for many Unicharm sites via the excat sitemap script (crawler fallback).
set +H 2>/dev/null || true
BASE="/backups/checksoundar/piramal-site-analysis/repo/unicharm-analysis/url-discovery"
PLUGIN_ROOT="/home/node/.excat-marketplaces/excat-marketplace/excat"
FETCH="${PLUGIN_ROOT}/skills/excat-url-discovery/scripts/fetch-sitemap.js"
CRAWL="${PLUGIN_ROOT}/skills/excat-url-discovery/scripts/crawl-site.js"
VALIDATE="${PLUGIN_ROOT}/tools/excatops-mcp/src/utils/validate-json-schema.js"
ASCHEMA="${PLUGIN_ROOT}/skills/excat-url-discovery/schemas/urls-all.schema.json"
GSCHEMA="${PLUGIN_ROOT}/skills/excat-url-discovery/schemas/urls-grouped.schema.json"
SAMPLE="${PLUGIN_ROOT}/skills/excat-url-discovery/scripts/create-sample.js"
RESULTS="${BASE}/batch-results.tsv"

process_json() {  # folder, method, sitemapPath
  local folder="$1"; local method="$2"; local smpath="$3"; local src="$4"
  cat "$src" | node -e '
    const fs=require("fs");
    const method=process.argv[1], smPath=process.argv[2], out=process.argv[3];
    let raw; try { raw=JSON.parse(fs.readFileSync(0,"utf8")); } catch(e){ raw=[]; }
    const docExt=[".pdf",".doc",".docx",".xls",".xlsx",".ppt",".pptx",".txt",".csv",".zip",".rar"];
    const imgExt=[".jpg",".jpeg",".png",".gif",".svg",".webp",".ico",".bmp",".tiff"];
    const urls=[],documents=[];
    raw.forEach(it=>{const u=typeof it==="string"?it:it.url;if(!u)return;const l=u.toLowerCase();
      const isDoc=docExt.some(e=>l.endsWith(e)), isImg=imgExt.some(e=>l.endsWith(e));
      const obj = (method==="sitemap")?{url:u}:{url:u,status:(typeof it==="string"?200:(it.status||200))};
      (isDoc&&!isImg)?documents.push(obj):urls.push(obj);});
    const conf = method==="sitemap"?"95%":method==="crawl"?"85%":"70%";
    const o={"analysis-urls-all":{captured:new Date().toISOString(),totalUrls:urls.length,totalDocuments:documents.length,
      method,sitemapURL:(method==="sitemap"?(smPath||"/sitemap.xml"):null),robotsTxtFound:method==="crawl",robotsTxtRulesApplied:method==="crawl",
      limitations:"",confidence:conf,urls,documents}};
    fs.writeFileSync(out,JSON.stringify(o,null,2));
  ' "$method" "$smpath" "$folder/urls-all.json"
}

run_site() {
  local key="$1"; local url="$2"
  local folder="$BASE/$key"
  mkdir -p "$folder"
  cat > "$folder/.catalog-config.json" << EOF
{"siteUrl":"$url","catalogFolder":"$folder","projectRoot":"/backups/checksoundar/piramal-site-analysis/repo","pluginRoot":"$PLUGIN_ROOT"}
EOF
  local method="" smpath="" cnt=0
  # ---- Priority 1: sitemap ----
  node "$FETCH" "$url" --logFile "$folder/catalog.log" > "$folder/.sm.tmp" 2>"$folder/.smerr.tmp"
  local ec=$?
  if [ $ec -eq 0 ]; then
    smpath=$(head -n1 "$folder/.sm.tmp" | sed -n 's/^SITEMAP_URL://p')
    tail -n +2 "$folder/.sm.tmp" > "$folder/.smjson.tmp"
    cnt=$(jq 'length' "$folder/.smjson.tmp" 2>/dev/null || echo 0)
    if [ "${cnt:-0}" -gt 0 ]; then method="sitemap"; fi
  fi
  # ---- Priority 2: crawl ----
  if [ -z "$method" ]; then
    node "$CRAWL" "$url" --max-pages 4000 --delay 300 --timeout 15000 --max-retries 2 \
      --checkpoint-file "$folder/crawl-checkpoint.json" --logFile "$folder/catalog.log" \
      > "$folder/.crawljson.tmp" 2>"$folder/.crawlerr.tmp"
    local cec=$?
    if [ $cec -eq 0 ] && jq empty "$folder/.crawljson.tmp" 2>/dev/null; then
      cnt=$(jq 'length' "$folder/.crawljson.tmp" 2>/dev/null || echo 0)
      if [ "${cnt:-0}" -ge 1 ]; then method="crawl"; fi
    fi
  fi

  if [ -z "$method" ]; then
    echo -e "${key}\t${url}\tFAILED\t0\t-\t$(head -c120 "$folder/.smerr.tmp" 2>/dev/null | tr '\n\t' '  ')" >> "$RESULTS"
    rm -f "$folder"/.sm.tmp "$folder"/.smerr.tmp "$folder"/.smjson.tmp "$folder"/.crawljson.tmp "$folder"/.crawlerr.tmp
    echo "[$key] FAILED"
    return
  fi

  local src="$folder/.smjson.tmp"; [ "$method" = "crawl" ] && src="$folder/.crawljson.tmp"
  process_json "$folder" "$method" "$smpath" "$src"
  node "$VALIDATE" "$folder/urls-all.json" "$ASCHEMA" >/dev/null 2>&1
  local av=$?
  node "$BASE/group.js" "$folder" "$url" "$key" >/dev/null 2>&1
  node "$VALIDATE" "$folder/urls-grouped.json" "$GSCHEMA" >/dev/null 2>&1
  node "$SAMPLE" "$folder" >/dev/null 2>&1
  jq -r '.["analysis-urls-all"].urls[].url' "$folder/urls-all.json" 2>/dev/null | sort > "$folder/urls-all.txt"
  local total=$(jq -r '.["analysis-urls-all"].totalUrls' "$folder/urls-all.json" 2>/dev/null)
  local vmark="OK"; [ $av -ne 0 ] && vmark="SCHEMA_WARN"
  echo -e "${key}\t${url}\t${vmark}\t${total}\t${method}\t${smpath}" >> "$RESULTS"
  rm -f "$folder"/.sm.tmp "$folder"/.smerr.tmp "$folder"/.smjson.tmp "$folder"/.crawljson.tmp "$folder"/.crawlerr.tmp
  echo "[$key] $vmark $method $total"
}

export -f run_site process_json
export BASE PLUGIN_ROOT FETCH CRAWL VALIDATE ASCHEMA GSCHEMA SAMPLE RESULTS

# site list: key<TAB>url  (passed on stdin)
while IFS=$'\t' read -r key url; do
  [ -z "$key" ] && continue
  run_site "$key" "$url"
done
