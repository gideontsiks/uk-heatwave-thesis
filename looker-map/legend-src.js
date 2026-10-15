// Heatwave Legend: a companion community visualization for the bubble map.
// Draws a horizontal white to cyan to red gradient bar and labels the low,
// middle, and high values of the same metric, so viewers can read the scale.
// Depends on the dscc global prepended at build time.

(function () {
  var el;

  function ensure() {
    if (!el) {
      el = document.createElement('div');
      el.id = 'legend';
      document.body.appendChild(el);
    }
    el.style.width = dscc.getWidth() + 'px';
    el.style.height = dscc.getHeight() + 'px';
  }

  function styleVal(style, id, fb) {
    if (!style || !style[id]) return fb;
    var v = (style[id].value !== undefined && style[id].value !== null)
      ? style[id].value : style[id].defaultValue;
    return (v === undefined || v === null || v === '') ? fb : v;
  }
  function colorVal(style, id, fb) {
    var v = styleVal(style, id, null);
    if (v === null) return fb;
    if (typeof v === 'object') return v.color || fb;
    return v;
  }
  function fmt(n, d) {
    if (isNaN(n)) return '';
    return Number(n).toFixed(d);
  }

  function draw(data) {
    ensure();
    var rows = (data.tables && data.tables.DEFAULT) || [];
    var c0 = colorVal(data.style, 'minColor', '#FFFFFF');
    var c1 = colorVal(data.style, 'midColor', '#0E7C8B');
    var c2 = colorVal(data.style, 'maxColor', '#C0392B');
    var title = styleVal(data.style, 'title', 'Excess per 100k');
    var dec = parseInt(styleVal(data.style, 'decimals', '1'), 10);
    if (isNaN(dec)) dec = 1;

    var vals = rows.map(function (r) { return Number(r.metric && r.metric[0]); })
      .filter(function (x) { return !isNaN(x); });
    var vmin = vals.length ? Math.min.apply(null, vals) : 0;
    var vmax = vals.length ? Math.max.apply(null, vals) : 1;
    var vmid = (vmin + vmax) / 2;

    el.innerHTML = '';
    var wrap = document.createElement('div');
    wrap.className = 'lg-wrap';

    if (title) {
      var t = document.createElement('div');
      t.className = 'lg-title';
      t.textContent = title;
      wrap.appendChild(t);
    }

    var bar = document.createElement('div');
    bar.className = 'lg-bar';
    bar.style.background = 'linear-gradient(to right, ' + c0 + ' 0%, ' + c1 + ' 50%, ' + c2 + ' 100%)';
    wrap.appendChild(bar);

    var labs = document.createElement('div');
    labs.className = 'lg-labs';
    labs.innerHTML = '<span>' + fmt(vmin, dec) + '</span>' +
      '<span>' + fmt(vmid, dec) + '</span>' +
      '<span>' + fmt(vmax, dec) + '</span>';
    wrap.appendChild(labs);

    el.appendChild(wrap);
  }

  dscc.subscribeToData(draw, { transform: dscc.objectTransform });
})();
