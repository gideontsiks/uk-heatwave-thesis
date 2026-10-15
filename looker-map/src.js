// Heatwave Bubble Map: a Looker Studio community visualization.
// Renders one bubble per row, positioned by Latitude and Longitude,
// sized and coloured by a single metric on a white to cyan to red scale.
// Depends on two globals that are prepended at build time: dscc and L (Leaflet).

(function () {
  var mapEl, map, tiles, layer;

  // Create the map container once, then keep it sized to the iframe.
  function ensureContainer() {
    if (!mapEl) {
      mapEl = document.createElement('div');
      mapEl.id = 'map';
      document.body.appendChild(mapEl);
    }
    mapEl.style.width = dscc.getWidth() + 'px';
    mapEl.style.height = dscc.getHeight() + 'px';
  }

  // Read a style value, falling back to the config default.
  function styleVal(style, id, fallback) {
    if (!style || !style[id]) return fallback;
    var v = (style[id].value !== undefined && style[id].value !== null)
      ? style[id].value : style[id].defaultValue;
    if (v === undefined || v === null || v === '') return fallback;
    return v;
  }

  // FILL_COLOR values can arrive as a string or as an object with a color field.
  function colorVal(style, id, fallback) {
    var v = styleVal(style, id, null);
    if (v === null) return fallback;
    if (typeof v === 'object') return v.color || fallback;
    return v;
  }

  function hx(c) {
    c = ('' + c).replace('#', '');
    return [parseInt(c.substr(0, 2), 16), parseInt(c.substr(2, 2), 16), parseInt(c.substr(4, 2), 16)];
  }
  function mix(a, b, t) {
    var A = hx(a), B = hx(b);
    return 'rgb(' + Math.round(A[0] + (B[0] - A[0]) * t) + ',' +
      Math.round(A[1] + (B[1] - A[1]) * t) + ',' +
      Math.round(A[2] + (B[2] - A[2]) * t) + ')';
  }
  // Two-segment ramp: low to mid for the bottom half, mid to high for the top.
  function ramp(t, c0, c1, c2) {
    t = Math.max(0, Math.min(1, t));
    return t < 0.5 ? mix(c0, c1, t * 2) : mix(c1, c2, (t - 0.5) * 2);
  }

  function drawViz(data) {
    ensureContainer();

    var rows = (data.tables && data.tables.DEFAULT) || [];
    var c0 = colorVal(data.style, 'minColor', '#FFFFFF');
    var c1 = colorVal(data.style, 'midColor', '#0E7C8B');
    var c2 = colorVal(data.style, 'maxColor', '#C0392B');
    var maxR = parseFloat(styleVal(data.style, 'maxRadius', '28')) || 28;
    var minR = 5;
    var opacity = parseFloat(styleVal(data.style, 'fillOpacity', '0.85'));
    if (isNaN(opacity)) opacity = 0.85;
    var showLabels = styleVal(data.style, 'showLabels', true);
    showLabels = (showLabels === true || showLabels === 'true');

    // Metric range across all rows, for normalising size and colour.
    var vals = rows.map(function (r) { return Number(r.metric && r.metric[0]); })
      .filter(function (x) { return !isNaN(x); });
    var vmin = vals.length ? Math.min.apply(null, vals) : 0;
    var vmax = vals.length ? Math.max.apply(null, vals) : 1;
    if (vmin === vmax) vmax = vmin + 1;

    if (!map) {
      map = L.map(mapEl, { scrollWheelZoom: true });
      tiles = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19, attribution: '&copy; OpenStreetMap contributors'
      }).addTo(map);
      layer = L.layerGroup().addTo(map);
    } else {
      map.invalidateSize();
      layer.clearLayers();
    }

    var pts = [];
    rows.forEach(function (r) {
      var lat = parseFloat(r.lat && r.lat[0]);
      var lon = parseFloat(r.lon && r.lon[0]);
      var v = Number(r.metric && r.metric[0]);
      if (isNaN(lat) || isNaN(lon) || isNaN(v)) return;

      var t = (v - vmin) / (vmax - vmin);
      var radius = minR + t * (maxR - minR);
      var fill = ramp(t, c0, c1, c2);
      var label = (r.label && r.label[0] != null) ? ('' + r.label[0]) : '';

      var cm = L.circleMarker([lat, lon], {
        radius: radius, color: '#33494D', weight: 1,
        fillColor: fill, fillOpacity: opacity
      });
      cm.bindPopup((label ? ('<b>' + label + '</b><br>') : '') + 'Value: ' + v);
      if (showLabels && label) {
        cm.bindTooltip(label, { permanent: true, direction: 'center', className: 'bub-label', opacity: 0.9 });
      }
      cm.addTo(layer);
      pts.push([lat, lon]);
    });

    if (pts.length) {
      map.fitBounds(pts, { padding: [30, 30], maxZoom: 7 });
    }
  }

  dscc.subscribeToData(drawViz, { transform: dscc.objectTransform });
})();
