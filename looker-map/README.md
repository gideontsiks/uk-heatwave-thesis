# Heatwave Region Map and Legend: custom Looker Studio community visualizations

This folder holds two custom charts for Looker Studio that work together:

- Heatwave Region Map: England and Wales NUTS1 regions drawn as a choropleth,
  each region shaded by a metric on a white to cyan to red scale, with region
  labels and a title. The region boundaries are embedded in the chart, so it
  needs no base-map tiles and no API key, and it always renders.
- Heatwave Legend: a horizontal white to cyan to red gradient bar that labels the
  low, middle, and high values of the same metric, so the map can be read at a
  glance.

A community visualization is just a set of files in a public Google Cloud Storage
bucket. Looker loads them by the bucket path. Both charts live in one bucket and
one manifest, so a single path gives you both.

## What is in this folder

Ready to upload (already built):
- `manifest.json` describes both charts and names their files.
- `heatmap-bubble.json`, `heatmap-bubble.js`, `heatmap-bubble.css` are the map.
- `heatmap-legend.json`, `heatmap-legend.js`, `heatmap-legend.css` are the legend.

Source, for rebuilding or editing:
- `src.js`, `custom.css` are the map source.
- `legend-src.js`, `legend-custom.css` are the legend source.
- `build.sh` fetches the libraries and rebuilds all four bundled files.

You can deploy the seven ready files as they are. You only need `build.sh` if you
edit any of the source files.

## Part 1: deploy to Google Cloud Storage

### Step 1: sign in to Google Cloud

You need a Google Cloud project and the gcloud tools. If you do not have them,
install the Google Cloud CLI from cloud.google.com/sdk, then run:

```
gcloud auth login && gcloud config set project YOUR_PROJECT_ID
```

Replace YOUR_PROJECT_ID with your project id from the Cloud console.

### Step 2: make a public bucket and upload the seven files

Pick a globally unique bucket name. From inside this folder, run this as one block,
replacing the bucket name with yours:

```
cd ~/Code/heatwave-map-viz && gsutil mb -b on gs://gtsikudo-heatwave-viz && gsutil iam ch allUsers:objectViewer gs://gtsikudo-heatwave-viz && gsutil cp manifest.json heatmap-bubble.json heatmap-bubble.js heatmap-bubble.css heatmap-legend.json heatmap-legend.js heatmap-legend.css gs://gtsikudo-heatwave-viz
```

What it does: `mb` makes the bucket, the `iam ch allUsers:objectViewer` line makes
its files publicly readable so Looker can fetch them, and `cp` uploads all seven
files. Your component path is now `gs://gtsikudo-heatwave-viz`, the bucket with no
file name on the end.

## Part 2: add both charts to your report

1. Open your report at lookerstudio.google.com and click Edit.
2. On the toolbar click "Community visualizations and components", then "Explore more".
3. Click "Build your own visualization" at the top.
4. In "Manifest path" paste your bucket path exactly: `gs://gtsikudo-heatwave-viz`
5. Click Submit. Two cards appear: Heatwave Region Map and Heatwave Legend. Click
   the map first to drop it on the canvas; you will add the legend in Part 4.

## Part 3: connect the map fields

Select the map, open the Setup tab, set the data source to `region_summary`, then:

- Region name goes to the Region field.
- Metric goes to ExcessPer100k.

That is all the map needs. It matches each row's Region name to the embedded
boundary of that region and shades it. The region names in your data must match
these exactly: North East, North West, Yorkshire and The Humber, East Midlands,
West Midlands, East of England, London, South East, South West, Wales. If one is
spelled differently in your sheet, that region will draw in grey as "no data", so
fix the spelling in the sheet to match.

## Part 4: add and connect the legend

1. With nothing selected, open "Community visualizations and components" again and
   this time click the Heatwave Legend card to drop it on the canvas.
2. Select the legend, open Setup, set the data source to `region_summary`, then:
   - Dimension goes to Region.
   - Metric goes to ExcessPer100k, the same metric the map uses.
3. The legend reads the low, middle, and high of that metric automatically, so its
   numbers always match the map.

## Part 5: lay it out so it reads as one unit

This is where the map becomes a panel a reviewer can read in seconds.

1. Size and place the map. Drag the map to fill the left two-thirds of the page,
   tall enough to show England and Wales without scrolling, for example 520 by 560
   pixels. Use the arrange handles, or set exact size and position in the Style
   tab's layout fields.
2. Put the legend directly under the map, not floating elsewhere. Make it as wide
   as the map and short, for example the map's width by 70 pixels. A legend that
   sits flush under the bar it describes reads as part of the same object.
3. Give the map a title above it with a text box, for example "Pooled excess
   heat-mortality by region, 2014 to 2022". Keep the text box the same width as the
   map so the three elements (title, map, legend) share one column.
4. Align all three. Select the title, the map, and the legend together, then use
   Arrange, Align, Left so their left edges line up, and Distribute if needed.

## Part 6: match the colours across both charts

Both charts default to the thesis scale, so they already match. To change it,
change it in both places so they stay in sync:

1. Select the map, open Style, and set Low to white, Mid to cyan (#0E7C8B), High to
   red (#C0392B). These are the defaults.
2. Select the legend, open Style, and set the same three colours.
3. On the map Style tab you can also set the map title, the fill opacity, whether
   region labels show, and the number of decimals. On the legend Style tab you can
   set the title text and the number of decimals.

## Part 7: final checks

- Click View, top right, to see the clean report. The reddest regions should be the
  Midlands and South East, with the North East near white, matching the thesis.
- Hover or click a region: a popup shows the region and its value. The labels on each
  region show its name and value.
- The legend's right-hand number should equal the largest regional value, and the
  left-hand number the smallest. If they do not, confirm both charts use the same
  metric and the same data source.

## Part 8: using the map as a filter

The map sends a filter on the Region field when you click a region, so it can
cross-filter other charts. In View mode, click a region to filter, click it again
to clear. The clicked region stays highlighted with a dark border; the others dim.

What it can filter: any chart whose data source has a matching Region field.

- Same data source as the map (region_summary): the regional bar chart, the H1
  greenspace scatter, the H2 building-density scatter, and the scorecards built on
  region_summary all respond. Click West Midlands and they show West Midlands only.
- Monthly time series (looker_monthly_mortality, which has a Region column): it can
  be filtered to the clicked region too, showing that region's monthly deaths and
  temperature. This works when its Region field name matches the map's.

What it should not filter, because these have no region dimension:

- The Moran's I by year chart and the DLNM exposure-response curve are national or
  temporal, so a region click cannot narrow them. Leave them unlinked.

Keeping the legend on the full scale: because cross-filtering applies to every chart
on the same data source, selecting one region would collapse the legend to a single
value. To keep the legend showing the whole low-to-high range, add the Google Sheet
tab a second time as a separate data source and point the legend at that copy, so the
map's filter does not reach it.

## While developing, and rebuilding

The manifest has `"devMode": true`, so Looker does not cache the files and your
edits show on reload. When the code is stable, set it to `"devMode": false` and
re-upload `manifest.json` for faster loading. If you edit any source file, rebuild
and re-upload:

```
cd ~/Code/heatwave-map-viz && bash build.sh && gsutil cp heatmap-bubble.js heatmap-bubble.css heatmap-legend.js heatmap-legend.css gs://gtsikudo-heatwave-viz
```

## Troubleshooting

- Chart will not load: the files are probably not public. Re-run the
  `gsutil iam ch allUsers:objectViewer` line, and check the manifest path ends at
  the bucket with no file name.
- A region draws grey as "no data": its name in your sheet does not match the
  boundary name. Fix the spelling to one of the ten listed in Part 3.
- Nothing shades at all: confirm Region is in the Region slot and the metric is set,
  and that the data source is region_summary.
- Legend numbers look wrong: it must use the same metric and data source as the map,
  with Region as its dimension.
- An error mentions dscId: that only happens when the files are opened outside
  Looker. Inside a report it does not occur.

## Cost and privacy

A tiny bucket with these files costs a few cents a month at most. The files are
public so Looker can read them, but they hold only visualization code, no data.
Your data stays in your Google Sheet and flows through Looker, never into the bucket.
