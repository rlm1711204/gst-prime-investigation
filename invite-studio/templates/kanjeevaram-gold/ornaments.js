/* Kanjeevaram Gold (KG-01) — ornaments drawn in code.
   Temple doors, mango-leaf thoranam, zari border, kolam corner, kuthuvilakku and jasmine petals.
   All artwork is original vector drawing; colours come from the theme's CSS variables. */
(function () {
  'use strict';
  var uid = 0;
  function cssVar(name, fallback) {
    var v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return v || fallback;
  }

  // One door leaf. The seam is on the right edge (x=200); the right door is the same SVG mirrored.
  function door() {
    var id = 'kgw' + (++uid);
    var studs = function (y) {
      var s = '';
      for (var x = 42; x <= 158; x += 29) s += '<circle cx="' + x + '" cy="' + y + '" r="4.6" class="kg-stud"/>';
      return s;
    };
    var panel = function (y, hgt) {
      var cx = 100, cy = y + hgt / 2;
      return '<rect x="30" y="' + y + '" width="140" height="' + hgt + '" rx="6" class="kg-panel"/>' +
        '<rect x="40" y="' + (y + 10) + '" width="120" height="' + (hgt - 20) + '" rx="3" class="kg-line"/>' +
        '<path d="M' + cx + ' ' + (cy - 38) + ' L' + (cx + 38) + ' ' + cy + ' L' + cx + ' ' + (cy + 38) + ' L' + (cx - 38) + ' ' + cy + 'Z" class="kg-line2"/>' +
        lotus(cx, cy, 15);
    };
    return '<svg viewBox="0 0 200 700" preserveAspectRatio="xMaxYMid slice" aria-hidden="true" focusable="false">' +
      '<defs><linearGradient id="' + id + '" x1="0" x2="1" y1="0" y2="0">' +
      '<stop offset="0" style="stop-color:var(--door-a)"/><stop offset=".6" style="stop-color:var(--door-b)"/><stop offset="1" style="stop-color:var(--door-a)"/>' +
      '</linearGradient></defs>' +
      '<rect width="200" height="700" fill="url(#' + id + ')"/>' +
      '<rect x="8" y="8" width="184" height="684" rx="4" class="kg-line" style="stroke-width:4"/>' +
      '<rect x="17" y="17" width="166" height="666" rx="3" class="kg-line"/>' +
      // arched top panel
      '<path d="M30 190V112Q30 42 100 42Q170 42 170 112V190Z" class="kg-panel"/>' +
      '<path d="M40 180V114Q40 54 100 54Q160 54 160 114V180Z" class="kg-line"/>' +
      lotus(100, 124, 22) +
      studs(203) + panel(216, 150) + studs(379) + panel(392, 140) + studs(545) + panel(558, 110) +
      // seam strip and knocker
      '<rect x="190" y="8" width="10" height="684" class="kg-seam"/>' +
      '<circle cx="176" cy="430" r="7" class="kg-stud"/>' +
      '<circle cx="176" cy="452" r="16" class="kg-ring"/>' +
      '</svg>';
  }

  function lotus(cx, cy, r) {
    var s = '<g class="kg-lotus">';
    for (var i = 0; i < 8; i++) {
      var a = i * 45;
      s += '<ellipse cx="' + cx + '" cy="' + (cy - r * 0.62) + '" rx="' + (r * 0.26) + '" ry="' + (r * 0.55) + '" transform="rotate(' + a + ' ' + cx + ' ' + cy + ')"/>';
    }
    return s + '<circle cx="' + cx + '" cy="' + cy + '" r="' + (r * 0.28) + '" class="kg-lotus-c"/></g>';
  }

  // Mango-leaf thoranam with marigolds, hung across the top of the doors.
  function thoranam() {
    var id = 'kgt' + (++uid);
    return '<svg width="100%" height="70" aria-hidden="true" focusable="false"><defs>' +
      '<pattern id="' + id + '" width="88" height="70" patternUnits="userSpaceOnUse">' +
      '<path d="M0 8Q22 20 44 8Q66 20 88 8" fill="none" style="stroke:var(--accent)" stroke-width="2"/>' +
      leaf(22, 14, '#2F6B1F') + leaf(66, 14, '#4C8C2B') +
      '<circle cx="0" cy="9" r="6.5" fill="#E8781A"/><circle cx="0" cy="9" r="3.2" fill="#FFB347"/>' +
      '<circle cx="44" cy="9" r="6.5" fill="#F2A31B"/><circle cx="44" cy="9" r="3.2" fill="#FFD27A"/>' +
      '<circle cx="88" cy="9" r="6.5" fill="#E8781A"/><circle cx="88" cy="9" r="3.2" fill="#FFB347"/>' +
      '</pattern></defs><rect width="100%" height="70" fill="url(#' + id + ')"/></svg>';
  }
  function leaf(x, y, c) {
    return '<path d="M' + x + ' ' + y + 'C' + (x + 11) + ' ' + (y + 12) + ' ' + (x + 10) + ' ' + (y + 34) + ' ' + x + ' ' + (y + 50) +
      'C' + (x - 10) + ' ' + (y + 34) + ' ' + (x - 11) + ' ' + (y + 12) + ' ' + x + ' ' + y + 'Z" fill="' + c + '"/>' +
      '<path d="M' + x + ' ' + (y + 3) + 'V' + (y + 46) + '" stroke="#CFE3B0" stroke-width=".8" opacity=".7"/>';
  }

  // Kuthuvilakku (brass oil lamp).
  function lamp() {
    return '<svg viewBox="0 0 60 120" aria-hidden="true" focusable="false" class="kg-lamp-svg">' +
      '<path class="kg-flame" d="M30 4C35 12 36 18 30 26C24 18 25 12 30 4Z" fill="#FFB300"/>' +
      '<path d="M30 12C32 16 32 19 30 23C28 19 28 16 30 12Z" fill="#FFF3C4"/>' +
      '<path d="M4 31Q30 46 56 31L52 29Q30 38 8 29Z" class="kg-brass"/>' +
      '<path d="M2 30L9 27L9 32ZM58 30L51 27L51 32Z" class="kg-brass"/>' +
      '<rect x="27" y="38" width="6" height="58" class="kg-brass"/>' +
      '<ellipse cx="30" cy="50" rx="7" ry="3.5" class="kg-brass"/>' +
      '<ellipse cx="30" cy="66" rx="8" ry="4" class="kg-brass"/>' +
      '<ellipse cx="30" cy="82" rx="7" ry="3.5" class="kg-brass"/>' +
      '<path d="M12 112Q30 92 48 112Z" class="kg-brass"/>' +
      '<rect x="8" y="111" width="44" height="6" rx="2" class="kg-brass"/>' +
      '</svg>';
  }

  // Pulli-kolam corner motif.
  function kolam() {
    var dots = '';
    [[14, 14], [36, 14], [14, 36], [36, 36], [58, 14], [14, 58]].forEach(function (d) {
      dots += '<circle cx="' + d[0] + '" cy="' + d[1] + '" r="2" class="kg-kolam-dot"/>';
    });
    return '<svg viewBox="0 0 80 80" aria-hidden="true" focusable="false"><g class="kg-kolam">' +
      '<path d="M3 76V30Q3 3 30 3H76"/>' +
      '<path d="M9 76V32Q9 9 32 9H76" opacity=".6"/>' +
      '<path d="M14 14Q30 18 36 36Q18 30 14 14Z"/>' +
      '<path d="M25 3C25 14 14 25 3 25"/>' +
      '<path d="M47 3C47 26 26 47 3 47" opacity=".75"/>' +
      '<path d="M36 14q11 0 11 11M14 36q0 11 11 11"/>' +
      '</g>' + dots + '</svg>';
  }

  // Repeating zari border as a CSS background (needs real colours, not CSS variables).
  function zariUrl() {
    var gold = cssVar('--accent', '#C9A227');
    var light = cssVar('--accent-light', '#E8D08A');
    var deep = cssVar('--zari-ink', '#3A070B');
    var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="30">' +
      '<rect y="0" width="40" height="1.5" fill="' + light + '"/>' +
      '<rect y="4" width="40" height="22" fill="' + gold + '"/>' +
      '<rect y="28.5" width="40" height="1.5" fill="' + light + '"/>' +
      '<circle cx="5" cy="8.5" r="1.4" fill="' + deep + '"/><circle cx="15" cy="8.5" r="1.4" fill="' + deep + '"/>' +
      '<circle cx="25" cy="8.5" r="1.4" fill="' + deep + '"/><circle cx="35" cy="8.5" r="1.4" fill="' + deep + '"/>' +
      '<path d="M0 26L10 12L20 26ZM20 26L30 12L40 26Z" fill="' + deep + '"/>' +
      '<path d="M6 26L10 19L14 26ZM26 26L30 19L34 26Z" fill="' + light + '"/>' +
      '</svg>';
    return 'url("data:image/svg+xml,' + encodeURIComponent(svg) + '")';
  }

  // Jasmine flowers, buds and a few marigold petals for the falling-petal canvas.
  function petalSprites(dpr) {
    function sprite(draw) {
      var c = document.createElement('canvas');
      var s = Math.round(24 * dpr);
      c.width = c.height = s;
      var x = c.getContext('2d');
      x.scale(s / 24, s / 24);
      draw(x);
      return c;
    }
    var jasmine = sprite(function (x) {
      x.translate(12, 12);
      x.fillStyle = '#FFFFFA';
      x.strokeStyle = 'rgba(160,140,90,.35)';
      x.lineWidth = 0.6;
      for (var i = 0; i < 5; i++) {
        x.rotate(Math.PI * 2 / 5);
        x.beginPath();
        x.ellipse(0, -6, 2.8, 6, 0, 0, Math.PI * 2);
        x.fill(); x.stroke();
      }
      x.fillStyle = '#F1DC8C';
      x.beginPath(); x.arc(0, 0, 1.8, 0, Math.PI * 2); x.fill();
    });
    var bud = sprite(function (x) {
      x.translate(12, 12);
      x.rotate(0.4);
      x.fillStyle = '#FFFDF2';
      x.beginPath(); x.ellipse(0, -2, 3.2, 8, 0, 0, Math.PI * 2); x.fill();
      x.fillStyle = '#7BA05B';
      x.beginPath(); x.ellipse(0, 7, 2, 2.6, 0, 0, Math.PI * 2); x.fill();
    });
    var petal = sprite(function (x) {
      x.translate(12, 12);
      x.fillStyle = '#FFF8E4';
      x.beginPath();
      x.moveTo(0, -9); x.quadraticCurveTo(7, -1, 0, 9); x.quadraticCurveTo(-5, -1, 0, -9);
      x.fill();
    });
    var marigold = sprite(function (x) {
      x.translate(12, 12);
      x.fillStyle = '#F29A1F';
      x.beginPath(); x.ellipse(0, 0, 4, 6.5, 0.5, 0, Math.PI * 2); x.fill();
    });
    return [jasmine, jasmine, bud, petal, petal, marigold];
  }

  function fill(sel, html) {
    Array.prototype.forEach.call(document.querySelectorAll(sel), function (e) { e.innerHTML = html; });
  }

  window.ORNAMENTS = {
    init: function () {
      fill('[data-orn="door"]', door());
      fill('[data-orn="thoranam"]', thoranam());
      fill('[data-orn="lamp"]', lamp());
      fill('[data-orn="kolam"]', kolam());
      document.documentElement.style.setProperty('--kg-zari', zariUrl());
    },
    petalSprites: petalSprites
  };
})();
