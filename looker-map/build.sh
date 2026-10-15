#!/bin/bash
# Builds the two bundled resource files the community visualization needs:
#   heatmap-bubble.js  = dscc helper library + Leaflet + your source
#   heatmap-bubble.css = Leaflet stylesheet + your custom styles
# Run this on your Mac, where the network is open. Needs node, npm, and curl.
set -e
cd "$(dirname "$0")"

echo "1. Fetching the dscc helper library from npm"
npm install @google/dscc >/dev/null 2>&1
cp node_modules/@google/dscc/_bundles/dscc.min.js .

echo "2. Fetching Leaflet (map library)"
curl -sL https://unpkg.com/leaflet@1.9.4/dist/leaflet.js  -o leaflet.js
curl -sL https://unpkg.com/leaflet@1.9.4/dist/leaflet.css -o leaflet.css

echo "3. Bundling the MAP (order matters: dscc, then Leaflet, then source)"
cat dscc.min.js leaflet.js src.js > heatmap-bubble.js
cat leaflet.css custom.css > heatmap-bubble.css

echo "4. Bundling the LEGEND (dscc + source; no map library needed)"
cat dscc.min.js legend-src.js > heatmap-legend.js
cp legend-custom.css heatmap-legend.css

echo "Done. Upload these to your bucket:"
echo "  manifest.json"
echo "  heatmap-bubble.json heatmap-bubble.js heatmap-bubble.css"
echo "  heatmap-legend.json heatmap-legend.js heatmap-legend.css"
