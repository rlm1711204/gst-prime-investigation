/* Invite engine — shared runtime used by every template.
   Reads window.INVITE (order data), window.BRAND, window.TEMPLATE and the
   optional window.ORNAMENTS hooks that a template's ornaments.js defines. */
(function () {
  'use strict';

  var D = window.INVITE;
  var B = window.BRAND || {};
  var T = window.TEMPLATE || {};
  var O = window.ORNAMENTS || {};
  var doc = document;
  var root = doc.documentElement;
  var body = doc.body;
  var reduceMotion = !!(window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches);
  var LANG_KEY = 'invite_lang';
  var lang = D.default_lang;
  var guest = '';
  var opened = false;

  root.classList.remove('no-js');

  /* ---------------- UI strings ---------------- */
  var UI = {
    tap_open: { ta: 'தொட்டுத் திறக்கவும்', en: 'Tap to open', hi: 'खोलने के लिए छुएँ' },
    dear: { ta: 'அன்புள்ள', en: 'Dear', hi: 'प्रिय' },
    weds: { ta: '&', en: 'weds', hi: 'संग' },
    and: { ta: '&', en: '&', hi: '&' },
    our_moments: { ta: 'எங்கள் தருணங்கள்', en: 'Our moments', hi: 'हमारे पल' },
    events: { ta: 'நிகழ்ச்சிகள்', en: 'Celebrations', hi: 'समारोह' },
    open_maps: { ta: 'வரைபடம்', en: 'Open in Maps', hi: 'मानचित्र' },
    add_cal: { ta: 'நாட்காட்டியில் சேர்', en: 'Add to calendar', hi: 'कैलेंडर में जोड़ें' },
    google_cal: { ta: 'கூகுள் நாட்காட்டி', en: 'Google Calendar', hi: 'गूगल कैलेंडर' },
    ics_cal: { ta: 'பிற நாட்காட்டி (.ics)', en: 'Other calendar (.ics)', hi: 'अन्य कैलेंडर (.ics)' },
    dress_code: { ta: 'உடை நிறம்', en: 'Dress code', hi: 'पोशाक' },
    countdown: { ta: 'இன்னும் சில நாட்களில்', en: 'The big day is almost here', hi: 'शुभ घड़ी नज़दीक है' },
    days: { ta: 'நாட்கள்', en: 'Days', hi: 'दिन' },
    hours: { ta: 'மணி', en: 'Hours', hi: 'घंटे' },
    mins: { ta: 'நிமிடம்', en: 'Mins', hi: 'मिनट' },
    secs: { ta: 'விநாடி', en: 'Secs', hi: 'सेकंड' },
    thanks_blessing: { ta: 'எங்களை வாழ்த்தியதற்கு நன்றி', en: 'Thank you for blessing us', hi: 'हमें आशीर्वाद देने के लिए धन्यवाद' },
    rsvp_title: { ta: 'தங்கள் வருகையை உறுதிசெய்யவும்', en: 'Will you join us?', hi: 'क्या आप पधारेंगे?' },
    rsvp_deadline: { ta: '{date}க்குள் பதில் அளிக்கவும்', en: 'Kindly reply by {date}', hi: 'कृपया {date} तक उत्तर दें' },
    rsvp_closed: { ta: 'பதில் அனுப்பும் நேரம் முடிந்தது', en: 'Replies are now closed', hi: 'उत्तर देने का समय समाप्त हो गया' },
    rsvp_preview_off: { ta: 'இது மாதிரிப் பக்கம் – பதில் அனுப்ப இயலாது.', en: 'This is a preview – replies are switched off.', hi: 'यह प्रीव्यू है – उत्तर भेजना बंद है।' },
    rsvp_not_ready: { ta: 'பதில் அனுப்பும் வசதி விரைவில்.', en: 'Replies will open soon.', hi: 'उत्तर भेजने की सुविधा जल्द शुरू होगी।' },
    your_reply: { ta: 'உங்கள் பதில்', en: 'Your reply', hi: 'आपका उत्तर' },
    attending: { ta: 'கட்டாயம் வருகிறோம்', en: 'Yes, attending', hi: 'हाँ, आएँगे' },
    not_attending: { ta: 'வர இயலாது', en: "Sorry, can't make it", hi: 'क्षमा करें, नहीं आ पाएँगे' },
    maybe: { ta: 'இன்னும் உறுதியில்லை', en: 'Maybe', hi: 'शायद' },
    guests_count: { ta: 'எத்தனை பேர்?', en: 'Number of guests', hi: 'कितने लोग?' },
    your_name: { ta: 'உங்கள் பெயர்', en: 'Your name', hi: 'आपका नाम' },
    your_wish: { ta: 'வாழ்த்துச் செய்தி (விருப்பம்)', en: 'Your wishes (optional)', hi: 'आपकी शुभकामनाएँ (वैकल्पिक)' },
    send: { ta: 'அனுப்பு', en: 'Send reply', hi: 'भेजें' },
    sending: { ta: 'அனுப்புகிறது…', en: 'Sending…', hi: 'भेजा जा रहा है…' },
    need_choice: { ta: 'தயவுசெய்து ஒரு பதிலைத் தேர்ந்தெடுக்கவும்.', en: 'Please choose a reply.', hi: 'कृपया एक उत्तर चुनें।' },
    need_name: { ta: 'தயவுசெய்து உங்கள் பெயரை எழுதவும்.', en: 'Please enter your name.', hi: 'कृपया अपना नाम लिखें।' },
    rsvp_thanks: { ta: 'நன்றி! உங்கள் பதில் கிடைத்தது.', en: 'Thank you! Your reply has been received.', hi: 'धन्यवाद! आपका उत्तर मिल गया।' },
    rsvp_again: { ta: 'மீண்டும் பதில் அனுப்ப', en: 'Send another reply', hi: 'फिर से उत्तर भेजें' },
    rsvp_error: { ta: 'அனுப்ப இயலவில்லை. இணைய இணைப்பைச் சரிபார்த்து மீண்டும் முயலவும்.', en: 'Could not send. Please check your internet and try again.', hi: 'भेजा नहीं जा सका। कृपया इंटरनेट जाँचें और फिर कोशिश करें।' },
    wishes_title: { ta: 'அன்பான வாழ்த்துகள்', en: 'Blessings from loved ones', hi: 'प्रियजनों की शुभकामनाएँ' },
    contact_title: { ta: 'தொடர்புக்கு', en: 'Contact the family', hi: 'परिवार से संपर्क करें' },
    call: { ta: 'அழைக்க', en: 'Call', hi: 'कॉल करें' },
    whatsapp: { ta: 'வாட்ஸ்அப்', en: 'WhatsApp', hi: 'व्हाट्सऐप' },
    wa_prefill: { ta: 'வணக்கம்! {names} அழைப்பிதழ் கிடைத்தது.', en: 'Namaste! We received the invitation for {names}.', hi: 'नमस्ते! हमें {names} का निमंत्रण मिला।' },
    share: { ta: 'பகிர்', en: 'Share', hi: 'शेयर करें' },
    share_text: { ta: 'அன்புடன் அழைக்கிறோம்! {title}', en: "You're invited! {title}", hi: 'आप सादर आमंत्रित हैं! {title}' },
    music_on: { ta: 'இசையை இயக்கு', en: 'Play music', hi: 'संगीत चालू करें' },
    music_off: { ta: 'இசையை நிறுத்து', en: 'Mute music', hi: 'संगीत बंद करें' },
    photo_n: { ta: 'புகைப்படம் {n} / {total}', en: 'Photo {n} of {total}', hi: 'फ़ोटो {n} / {total}' },
    close: { ta: 'மூடு', en: 'Close', hi: 'बंद करें' },
    prev: { ta: 'முந்தையது', en: 'Previous', hi: 'पिछला' },
    next: { ta: 'அடுத்தது', en: 'Next', hi: 'अगला' },
    language: { ta: 'மொழி', en: 'Language', hi: 'भाषा' },
    preview_ribbon: { ta: 'மாதிரி – இறுதி வடிவம் அல்ல', en: 'PREVIEW – not the final invitation', hi: 'प्रीव्यू – अंतिम निमंत्रण नहीं' }
  };
  var LANG_LABEL = { ta: 'தமிழ்', en: 'EN', hi: 'हिं' };

  var MONTHS = {
    en: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    ta: ['ஜனவரி', 'பிப்ரவரி', 'மார்ச்', 'ஏப்ரல்', 'மே', 'ஜூன்', 'ஜூலை', 'ஆகஸ்ட்', 'செப்டம்பர்', 'அக்டோபர்', 'நவம்பர்', 'டிசம்பர்'],
    hi: ['जनवरी', 'फ़रवरी', 'मार्च', 'अप्रैल', 'मई', 'जून', 'जुलाई', 'अगस्त', 'सितंबर', 'अक्टूबर', 'नवंबर', 'दिसंबर']
  };
  var WEEKDAYS = {
    en: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    ta: ['ஞாயிற்றுக்கிழமை', 'திங்கட்கிழமை', 'செவ்வாய்க்கிழமை', 'புதன்கிழமை', 'வியாழக்கிழமை', 'வெள்ளிக்கிழமை', 'சனிக்கிழமை'],
    hi: ['रविवार', 'सोमवार', 'मंगलवार', 'बुधवार', 'गुरुवार', 'शुक्रवार', 'शनिवार']
  };

  /* ---------------- Helpers ---------------- */
  function $(sel, ctx) { return (ctx || doc).querySelector(sel); }
  function $$(sel, ctx) { return Array.prototype.slice.call((ctx || doc).querySelectorAll(sel)); }

  // Returns [text, languageActuallyUsed] with fallback lang → en → ta → hi.
  function pick(obj, l) {
    if (obj == null) return ['', l];
    if (typeof obj === 'string') return [obj, l];
    var order = [l || lang, 'en', 'ta', 'hi'];
    for (var i = 0; i < order.length; i++) {
      if (obj[order[i]]) return [obj[order[i]], order[i]];
    }
    return ['', l];
  }
  function tr(obj, l) { return pick(obj, l)[0]; }
  function ui(key, vars) {
    var s = tr(UI[key]);
    if (vars) s = s.replace(/\{(\w+)\}/g, function (m, k) { return vars[k] != null ? vars[k] : m; });
    return s;
  }
  function get(obj, path) {
    return path.split('.').reduce(function (o, k) { return o == null ? undefined : o[k]; }, obj);
  }
  // Builds an element. `html` is only ever used for icon/ornament SVG strings defined in code.
  function h(tag, attrs, kids) {
    var e = doc.createElement(tag);
    if (attrs) {
      for (var k in attrs) {
        if (!Object.prototype.hasOwnProperty.call(attrs, k) || attrs[k] == null || attrs[k] === false) continue;
        if (k === 'text') e.textContent = attrs[k];
        else if (k === 'html') e.innerHTML = attrs[k];
        else if (k === 'class') e.className = attrs[k];
        else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), attrs[k]);
        else e.setAttribute(k, attrs[k] === true ? '' : attrs[k]);
      }
    }
    (kids || []).forEach(function (c) { if (c) e.appendChild(typeof c === 'string' ? doc.createTextNode(c) : c); });
    return e;
  }
  function lsGet(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { window.localStorage.setItem(k, v); } catch (e) { /* private mode */ } }
  function digits(n) { return String(n || '').replace(/\D/g, ''); }
  function isRealNumber(n) { return !!n && !/x/i.test(n) && digits(n).length >= 10; }
  function safe(fn) { try { fn(); } catch (e) { if (window.console) console.error('[invite]', e); } }

  var ICON = {
    cal: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4.5" width="18" height="16" rx="2"/><path d="M3 9.5h18M8 2.5v4M16 2.5v4"/></svg>',
    clock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
    pin: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" aria-hidden="true"><path d="M12 21s-7-6.2-7-11.5A7 7 0 0 1 19 9.5C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/></svg>',
    phone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" aria-hidden="true"><path d="M5 3h4l2 5-2.5 1.5a11 11 0 0 0 6 6L16 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 5a2 2 0 0 1 2-2z"/></svg>',
    chat: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" aria-hidden="true"><path d="M4 20l1.4-4.2A8 8 0 1 1 8.5 19z"/><path d="M9 10h6M9 13.5h4"/></svg>',
    share: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="18" cy="5" r="2.6"/><circle cx="6" cy="12" r="2.6"/><circle cx="18" cy="19" r="2.6"/><path d="M8.3 10.8l7.4-4.3M8.3 13.2l7.4 4.3"/></svg>',
    soundOn: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 9h4l5-4v14l-5-4H4z"/><path d="M16.5 8.5a5 5 0 0 1 0 7M19 6a8.5 8.5 0 0 1 0 12"/></svg>',
    soundOff: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 9h4l5-4v14l-5-4H4z"/><path d="M17 9l5 6M22 9l-5 6"/></svg>',
    dress: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round" aria-hidden="true"><path d="M9 3l3 3 3-3 2 4-2 2 4 12H5l4-12-2-2z"/></svg>',
    tap: '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 11V5.5a1.5 1.5 0 0 1 3 0V11m0-1.5a1.5 1.5 0 0 1 3 0V11m0-.5a1.5 1.5 0 0 1 3 0V15a6 6 0 0 1-6 6h-.6a6 6 0 0 1-4.6-2.2L4.5 15.5a1.5 1.5 0 0 1 2.3-1.9L9 16"/></svg>'
  };

  /* ---------------- Dates (shown as written in the data, i.e. venue time) ---------------- */
  function parts(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(iso || '');
    if (!m) return null;
    var p = { y: +m[1], mo: +m[2], d: +m[3], h: +m[4], mi: +m[5] };
    p.wd = new Date(Date.UTC(p.y, p.mo - 1, p.d)).getUTCDay();
    return p;
  }
  function fmtDate(p, l) {
    l = l || lang;
    var mon = MONTHS[l] || MONTHS.en, wd = WEEKDAYS[l] || WEEKDAYS.en;
    if (l === 'ta') return p.d + ' ' + mon[p.mo - 1] + ' ' + p.y + ', ' + wd[p.wd];
    return wd[p.wd] + ', ' + p.d + ' ' + mon[p.mo - 1] + ' ' + p.y;
  }
  function fmtShortDate(p, l) {
    l = l || lang;
    var mon = MONTHS[l] || MONTHS.en;
    return p.d + ' ' + (l === 'en' ? mon[p.mo - 1].slice(0, 3) : mon[p.mo - 1]) + ' ' + p.y;
  }
  function fmtTime(p, l) {
    l = l || lang;
    var h12 = p.h % 12 || 12, mm = (p.mi < 10 ? '0' : '') + p.mi;
    if (l === 'en') return h12 + ':' + mm + (p.h < 12 ? ' AM' : ' PM');
    var periods = l === 'ta' ? ['இரவு', 'காலை', 'மதியம்', 'மாலை', 'இரவு'] : ['रात', 'सुबह', 'दोपहर', 'शाम', 'रात'];
    var i = p.h < 4 ? 0 : p.h < 12 ? 1 : p.h < 16 ? 2 : p.h < 19 ? 3 : 4;
    return periods[i] + ' ' + h12 + ':' + mm;
  }

  /* ---------------- Guest name from ?to= ---------------- */
  function readGuest() {
    if (D.guest_personalisation === false) return '';
    var raw = null;
    try { raw = new URLSearchParams(location.search).get('to'); } catch (e) { return ''; }
    if (!raw) return '';
    var s = raw.replace(/[-_+.]+/g, ' ');
    try { s = s.replace(new RegExp('[^\\p{L}\\p{M} ]', 'gu'), ''); }
    catch (e) { s = s.replace(/[^A-Za-z\u0B80-\u0BFF\u0900-\u097F ]/g, ''); }
    s = s.replace(/\s+/g, ' ').trim();
    var chars = Array.from ? Array.from(s) : s.split('');
    return chars.slice(0, 40).join('').trim();
  }

  /* ---------------- Names ---------------- */
  function namesText(l) {
    var a = tr(D.hosts.primary, l), b = tr(D.hosts.secondary, l);
    return b ? a + ' & ' + b : a;
  }

  /* ---------------- Static text bindings ---------------- */
  function bindText() {
    $$('[data-t]').forEach(function (e) {
      var r = pick(get(D, e.getAttribute('data-t')));
      e.textContent = r[0];
      e.setAttribute('lang', r[1]);
      e.hidden = !r[0];
    });
    $$('[data-ui]').forEach(function (e) { e.textContent = ui(e.getAttribute('data-ui')); });
    $$('[data-ui-aria]').forEach(function (e) { e.setAttribute('aria-label', ui(e.getAttribute('data-ui-aria'))); });
    $$('[data-ui-ph]').forEach(function (e) { e.setAttribute('placeholder', ui(e.getAttribute('data-ui-ph'))); });
    var hasSecond = !!tr(D.hosts.secondary);
    $$('[data-joiner]').forEach(function (e) {
      var kind = e.getAttribute('data-joiner');
      e.textContent = kind === 'amp' ? '&' : (tr(D.hosts.joiner) || ui('weds'));
      e.hidden = !hasSecond;
    });
    $$('[data-guest]').forEach(function (e) {
      e.textContent = guest ? ui('dear') + ' ' + guest + ',' : '';
      e.hidden = !guest;
    });
  }

  /* ---------------- Language ---------------- */
  function renderLangToggle() {
    var nav = $('#lang-toggle');
    if (!nav) return;
    nav.innerHTML = '';
    if (D.langs.length < 2) { nav.hidden = true; return; }
    nav.setAttribute('aria-label', ui('language'));
    D.langs.forEach(function (l) {
      nav.appendChild(h('button', {
        type: 'button', lang: l, 'aria-pressed': String(l === lang), text: LANG_LABEL[l],
        onclick: function () { setLang(l); }
      }));
    });
  }
  function setLang(l) {
    if (D.langs.indexOf(l) < 0) return;
    lang = l;
    root.setAttribute('lang', l);
    lsSet(LANG_KEY, l);
    renderAll();
  }

  /* ---------------- Photos + lightbox ---------------- */
  var photos = (D.photos || []).map(function (p) { return typeof p === 'string' ? { src: p } : p; });
  function photoAlt(i) {
    return tr(photos[i].alt) || ui('photo_n', { n: i + 1, total: photos.length }) + ' – ' + namesText();
  }
  function renderPhotos() {
    var wrap = $('#photo-stack');
    var scene = $('[data-scene="photos"]');
    if (!wrap) return;
    if (!photos.length) { if (scene) scene.hidden = true; return; }
    if (!wrap.children.length) {
      wrap.classList.toggle('single', photos.length === 1);
      photos.forEach(function (p, i) {
        var btn = h('button', { type: 'button', class: 'polaroid', 'data-i': i, onclick: function () { openLightbox(i); } }, [
          h('img', { src: p.src, alt: '', loading: i < 2 ? 'eager' : 'lazy', decoding: 'async', width: 1080, height: 1350 })
        ]);
        btn.style.transitionDelay = (i % 2) * 0.18 + 's';
        wrap.appendChild(btn);
      });
    }
    $$('.polaroid', wrap).forEach(function (b, i) {
      $('img', b).alt = photoAlt(i);
      b.setAttribute('aria-label', photoAlt(i));
    });
  }
  var lb = null, lbIndex = 0, lbReturn = null;
  function openLightbox(i) {
    lbReturn = doc.activeElement;
    lbIndex = i;
    var img = h('img', { alt: '' });
    var count = h('p', { class: 'lb-count', 'aria-live': 'polite' });
    lb = h('div', { class: 'lightbox', role: 'dialog', 'aria-modal': 'true' }, [
      img,
      h('button', { type: 'button', class: 'lb-btn lb-close', 'aria-label': ui('close'), text: '×', onclick: closeLightbox }),
      photos.length > 1 ? h('button', { type: 'button', class: 'lb-btn lb-prev', 'aria-label': ui('prev'), text: '‹', onclick: function () { stepLightbox(-1); } }) : null,
      photos.length > 1 ? h('button', { type: 'button', class: 'lb-btn lb-next', 'aria-label': ui('next'), text: '›', onclick: function () { stepLightbox(1); } }) : null,
      count
    ]);
    var x0 = null;
    lb.addEventListener('touchstart', function (e) { x0 = e.touches[0].clientX; }, { passive: true });
    lb.addEventListener('touchend', function (e) {
      if (x0 == null) return;
      var dx = e.changedTouches[0].clientX - x0;
      if (Math.abs(dx) > 45) stepLightbox(dx < 0 ? 1 : -1);
      x0 = null;
    });
    lb.addEventListener('click', function (e) { if (e.target === lb) closeLightbox(); });
    body.appendChild(lb);
    body.classList.add('is-locked');
    showLightbox();
    $('.lb-close', lb).focus();
  }
  function showLightbox() {
    var img = $('img', lb);
    img.src = photos[lbIndex].src;
    img.alt = photoAlt(lbIndex);
    $('.lb-count', lb).textContent = ui('photo_n', { n: lbIndex + 1, total: photos.length });
  }
  function stepLightbox(d) {
    lbIndex = (lbIndex + d + photos.length) % photos.length;
    showLightbox();
  }
  function closeLightbox() {
    if (!lb) return;
    lb.remove();
    lb = null;
    body.classList.remove('is-locked');
    if (lbReturn && lbReturn.focus) lbReturn.focus();
  }
  doc.addEventListener('keydown', function (e) {
    if (!lb) return;
    if (e.key === 'Escape') closeLightbox();
    else if (e.key === 'ArrowRight') stepLightbox(1);
    else if (e.key === 'ArrowLeft') stepLightbox(-1);
  });

  /* ---------------- Events ---------------- */
  function icsStamp(ms) { return new Date(ms).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, ''); }
  function icsEsc(s) { return String(s).replace(/\\/g, '\\\\').replace(/\r?\n/g, '\\n').replace(/([,;])/g, '\\$1'); }
  function eventTimes(ev) {
    var s = Date.parse(ev.start);
    var e = ev.end ? Date.parse(ev.end) : s + 2 * 3600 * 1000;
    return [s, e];
  }
  function calTitle(ev) { return tr(ev.title) + ' – ' + namesText(); }
  function calLocation(ev) { return [tr(ev.venue.name), tr(ev.venue.address)].filter(Boolean).join(', '); }
  function icsHref(ev) {
    var t = eventTimes(ev);
    var lines = [
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Invite//EN', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
      'BEGIN:VEVENT',
      'UID:' + D.slug + '-' + ev.id + '@' + (B.domain || 'invite'),
      'DTSTAMP:' + icsStamp(Date.now()),
      'DTSTART:' + icsStamp(t[0]),
      'DTEND:' + icsStamp(t[1]),
      'SUMMARY:' + icsEsc(calTitle(ev)),
      'LOCATION:' + icsEsc(calLocation(ev)),
      'DESCRIPTION:' + icsEsc(location.origin + location.pathname),
      'END:VEVENT', 'END:VCALENDAR'
    ];
    return 'data:text/calendar;charset=utf-8,' + encodeURIComponent(lines.join('\r\n'));
  }
  function gcalHref(ev) {
    var t = eventTimes(ev);
    return 'https://calendar.google.com/calendar/render?action=TEMPLATE' +
      '&text=' + encodeURIComponent(calTitle(ev)) +
      '&dates=' + icsStamp(t[0]) + '/' + icsStamp(t[1]) +
      '&location=' + encodeURIComponent(calLocation(ev)) +
      '&details=' + encodeURIComponent(location.origin + location.pathname);
  }
  function row(icon, kids) {
    return h('div', { class: 'event-row' }, [h('span', { class: 'ico', html: icon }), h('div', null, kids)]);
  }
  function renderEvents() {
    var list = $('#events-list');
    if (!list) return;
    list.innerHTML = '';
    D.events.forEach(function (ev, i) {
      var s = parts(ev.start), e = ev.end ? parts(ev.end) : null;
      var titleR = pick(ev.title);
      var card = h('article', { class: 'event-card reveal', 'data-delay': String(Math.min(i, 3)) });
      card.appendChild(h('h3', { class: 'event-title', lang: titleR[1], text: titleR[0] }));
      var dateKids = [h('strong', { text: fmtDate(s) })];
      if (ev.traditional_date) dateKids.push(h('div', { class: 'event-trad', text: tr(ev.traditional_date) }));
      card.appendChild(row(ICON.cal, dateKids));
      card.appendChild(row(ICON.clock, [fmtTime(s) + (e ? ' – ' + fmtTime(e) : '')]));
      var venueKids = [h('strong', { text: tr(ev.venue.name) })];
      if (ev.venue.address) venueKids.push(h('div', { class: 'venue-addr', text: tr(ev.venue.address) }));
      card.appendChild(row(ICON.pin, venueKids));
      if (ev.note) card.appendChild(h('p', { class: 'venue-addr', text: tr(ev.note) }));
      if (ev.dress_code) {
        var chip = h('span', { class: 'dress-chip' }, [h('i', { 'aria-hidden': 'true' }), ui('dress_code') + ': ' + tr(ev.dress_code.label)]);
        if (ev.dress_code.colour) $('i', chip).style.background = ev.dress_code.colour;
        else $('i', chip).remove();
        card.appendChild(chip);
      }
      var calMenu = h('div', { class: 'cal-menu', hidden: true }, [
        h('a', { class: 'btn', href: gcalHref(ev), target: '_blank', rel: 'noopener', text: ui('google_cal') }),
        h('a', { class: 'btn', href: icsHref(ev), download: D.slug + '-' + ev.id + '.ics', text: ui('ics_cal') })
      ]);
      var actions = h('div', { class: 'event-actions' });
      if (ev.venue.maps_url) {
        actions.appendChild(h('a', { class: 'btn btn-solid', href: ev.venue.maps_url, target: '_blank', rel: 'noopener', html: ICON.pin }, [ui('open_maps')]));
      }
      actions.appendChild(h('button', {
        type: 'button', class: 'btn', 'aria-expanded': 'false', html: ICON.cal,
        onclick: function () {
          calMenu.hidden = !calMenu.hidden;
          this.setAttribute('aria-expanded', String(!calMenu.hidden));
        }
      }, [ui('add_cal')]));
      actions.appendChild(calMenu);
      card.appendChild(actions);
      list.appendChild(card);
    });
    observeReveals(list);
  }

  /* ---------------- Countdown ---------------- */
  var cdTimer = null;
  function countdownEvent() {
    var id = D.countdown_to;
    for (var i = 0; i < D.events.length; i++) if (D.events[i].id === id) return D.events[i];
    return D.events[0];
  }
  function renderCountdown() {
    var box = $('#countdown');
    if (!box) return;
    var ev = countdownEvent();
    var target = Date.parse(ev.start);
    box.innerHTML = '';
    var cells = {};
    var grid = h('div', { class: 'countdown', role: 'timer' });
    ['days', 'hours', 'mins', 'secs'].forEach(function (k) {
      cells[k] = h('span', { class: 'cd-num', text: '0' });
      grid.appendChild(h('div', { class: 'cd-cell' }, [cells[k], h('span', { class: 'cd-lbl', text: ui(k) })]));
    });
    var done = h('p', { class: 'cd-done', text: ui('thanks_blessing'), hidden: true });
    var titleEl = $('#countdown-title');
    box.appendChild(grid);
    box.appendChild(done);
    box.appendChild(h('p', { class: 'cd-event', text: tr(ev.title) + ' · ' + fmtShortDate(parts(ev.start)) }));
    function tick() {
      var ms = target - Date.now();
      if (ms <= 0) {
        grid.hidden = true; done.hidden = false;
        if (titleEl) titleEl.hidden = true;
        clearInterval(cdTimer);
        return;
      }
      var s = Math.floor(ms / 1000);
      cells.days.textContent = Math.floor(s / 86400);
      cells.hours.textContent = pad(Math.floor(s / 3600) % 24);
      cells.mins.textContent = pad(Math.floor(s / 60) % 60);
      cells.secs.textContent = pad(s % 60);
    }
    function pad(n) { return (n < 10 ? '0' : '') + n; }
    clearInterval(cdTimer);
    tick();
    cdTimer = setInterval(tick, 1000);
  }

  /* ---------------- RSVP + wishes ---------------- */
  var rsvp = D.rsvp || { enabled: false };
  var endpoint = rsvp.endpoint || '';
  var endpointReady = /^https:\/\//.test(endpoint) && endpoint.indexOf('XXXX') < 0;
  function deadlinePassed() {
    if (!rsvp.deadline) return false;
    return Date.now() > Date.parse(rsvp.deadline + 'T23:59:59+05:30');
  }
  function buildRsvp() {
    var scene = $('[data-scene="rsvp"]');
    var host = $('#rsvp-root');
    if (!host) return;
    if (!rsvp.enabled) { if (scene) scene.hidden = true; return; }
    var maxG = rsvp.max_guests || 10;
    var sel = h('select', { id: 'rsvp-count', name: 'count' });
    for (var n = 1; n <= maxG; n++) sel.appendChild(h('option', { value: n, text: String(n) }));
    function choice(val, key) {
      return h('label', { class: 'choice' }, [
        h('input', { type: 'radio', name: 'status', value: val }),
        h('span', { 'data-ui': key })
      ]);
    }
    var note = h('p', { class: 'form-note', id: 'rsvp-note', 'aria-live': 'polite' });
    var submit = h('button', { type: 'submit', class: 'btn btn-solid', 'data-ui': 'send' });
    var countWrap = h('div', null, [h('label', { for: 'rsvp-count', 'data-ui': 'guests_count' }), sel]);
    var form = h('form', { class: 'rsvp-card', id: 'rsvp-form', novalidate: true }, [
      h('fieldset', null, [
        h('legend', { 'data-ui': 'your_reply' }),
        h('div', { class: 'choice-row' }, [choice('attending', 'attending'), choice('not_attending', 'not_attending'), choice('maybe', 'maybe')])
      ]),
      countWrap,
      h('label', { for: 'rsvp-name', 'data-ui': 'your_name' }),
      h('input', { type: 'text', id: 'rsvp-name', name: 'name', maxlength: '60', autocomplete: 'name', value: guest || null }),
      h('label', { for: 'rsvp-msg', 'data-ui': 'your_wish' }),
      h('textarea', { id: 'rsvp-msg', name: 'message', maxlength: '300', rows: '3' }),
      submit,
      note
    ]);
    form.addEventListener('change', function (e) {
      if (e.target.name !== 'status') return;
      $$('.choice', form).forEach(function (c) { c.classList.toggle('is-checked', $('input', c).checked); });
      countWrap.hidden = e.target.value === 'not_attending';
    });
    var thanks = h('div', { class: 'rsvp-card rsvp-thanks', hidden: true }, [
      h('p', { 'data-ui': 'rsvp_thanks' }),
      h('button', {
        type: 'button', class: 'btn', 'data-ui': 'rsvp_again',
        onclick: function () { thanks.hidden = true; form.hidden = false; }
      })
    ]);
    host.appendChild(form);
    host.appendChild(thanks);

    var blockKey = D.watermark ? 'rsvp_preview_off' : !endpointReady ? 'rsvp_not_ready' : deadlinePassed() ? 'rsvp_closed' : '';
    if (blockKey) {
      submit.disabled = true;
      note.setAttribute('data-ui', blockKey);
    }
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (blockKey) return;
      var status = (form.querySelector('input[name="status"]:checked') || {}).value;
      var name = $('#rsvp-name').value.replace(/\s+/g, ' ').trim();
      note.classList.add('err');
      if (!status) { note.textContent = ui('need_choice'); return; }
      if (name.length < 2) { note.textContent = ui('need_name'); $('#rsvp-name').focus(); return; }
      note.classList.remove('err');
      note.textContent = ui('sending');
      submit.disabled = true;
      var data = new URLSearchParams();
      data.set('slug', D.slug);
      data.set('name', name);
      data.set('status', status);
      data.set('count', status === 'not_attending' ? '0' : sel.value);
      data.set('message', $('#rsvp-msg').value.trim().slice(0, 300));
      data.set('guest_param', guest);
      data.set('lang', lang);
      data.set('timestamp', new Date().toISOString());
      // Apps Script does not send CORS headers on POST, so this is a no-cors "simple" request:
      // the response is opaque and a resolved promise means the request reached Google.
      fetch(endpoint, { method: 'POST', mode: 'no-cors', body: data }).then(function () {
        submit.disabled = false;
        note.textContent = '';
        form.hidden = true;
        thanks.hidden = false;
        lsSet('rsvp_' + D.slug, status);
        var msg = data.get('message');
        if (msg) prependWish({ name: name, message: msg });
        setTimeout(loadWishes, 4000);
      }).catch(function () {
        submit.disabled = false;
        note.classList.add('err');
        note.textContent = ui('rsvp_error');
      });
    });
  }
  function renderRsvpText() {
    var dl = $('#rsvp-deadline');
    if (dl) {
      var p = rsvp.deadline ? parts(rsvp.deadline + 'T00:00') : null;
      dl.textContent = p ? ui('rsvp_deadline', { date: fmtShortDate(p) }) : '';
      dl.hidden = !p || deadlinePassed();
    }
  }
  function wishEl(w) {
    return h('div', { class: 'wish' }, [h('p', { text: String(w.message || '').slice(0, 300) }), h('b', { text: '— ' + String(w.name || '').slice(0, 60) })]);
  }
  function prependWish(w) {
    var list = $('#wishes-list');
    if (!list) return;
    list.insertBefore(wishEl(w), list.firstChild);
    $('#wishes').hidden = false;
  }
  function loadWishes() {
    var list = $('#wishes-list');
    if (!list || !rsvp.enabled || !endpointReady || !window.fetch) return;
    var url = endpoint + (endpoint.indexOf('?') < 0 ? '?' : '&') + 'slug=' + encodeURIComponent(D.slug);
    fetch(url).then(function (r) { return r.json(); }).then(function (j) {
      var items = (j && j.wishes) || [];
      list.innerHTML = '';
      items.slice(0, 10).forEach(function (w) { list.appendChild(wishEl(w)); });
      $('#wishes').hidden = !items.length;
    }).catch(function () { /* wishes are optional */ });
  }

  /* ---------------- Contacts ---------------- */
  function renderContacts() {
    var list = $('#contacts-list');
    var scene = $('[data-scene="contacts"]');
    if (!list) return;
    var contacts = D.contacts || [];
    if (!contacts.length) { if (scene) scene.hidden = true; return; }
    list.innerHTML = '';
    contacts.forEach(function (c) {
      var actions = h('div', { class: 'contact-actions' });
      if (isRealNumber(c.phone)) actions.appendChild(h('a', { class: 'btn', href: 'tel:+' + digits(c.phone), html: ICON.phone }, [ui('call')]));
      if (isRealNumber(c.whatsapp)) {
        actions.appendChild(h('a', {
          class: 'btn btn-solid', target: '_blank', rel: 'noopener', html: ICON.chat,
          href: 'https://wa.me/' + digits(c.whatsapp) + '?text=' + encodeURIComponent(ui('wa_prefill', { names: namesText() }))
        }, [ui('whatsapp')]));
      }
      list.appendChild(h('div', { class: 'contact-card reveal' }, [h('h3', { text: tr(c.label) }), actions]));
    });
    observeReveals(list);
  }

  /* ---------------- Footer ---------------- */
  function renderFooter() {
    var f = $('#site-footer');
    if (!f) return;
    var brand = tr(B.name_i18n) || B.name || '';
    var code = T.code || '';
    var ft = B.footer || {};
    f.innerHTML = '';
    f.appendChild(h('p', { text: (tr(ft.by) || 'Invitation by') + ' ' + brand }));
    if (isRealNumber(B.whatsapp)) {
      var prefill = tr(ft.prefill).replace('{brand}', B.name || brand).replace('{code}', code);
      f.appendChild(h('p', null, [
        (tr(ft.want_one) || 'Want one? WhatsApp') + ' ',
        h('a', { href: 'https://wa.me/' + digits(B.whatsapp) + '?text=' + encodeURIComponent(prefill), target: '_blank', rel: 'noopener', text: B.whatsapp })
      ]));
    }
  }

  /* ---------------- Share + music buttons ---------------- */
  function shareUrl() { return location.origin + location.pathname; }
  function doShare() {
    var title = (D.og && D.og.title) || namesText();
    var text = ui('share_text', { title: title });
    if (navigator.share) {
      navigator.share({ title: title, text: text, url: shareUrl() }).catch(function () { /* cancelled */ });
    } else {
      window.open('https://wa.me/?text=' + encodeURIComponent(text + '\n' + shareUrl()), '_blank', 'noopener');
    }
  }

  var audio = null, userMuted = false, fadeTimer = null;
  function setupAudio() {
    var btn = $('#fab-mute');
    if (!D.music || !D.music.file) { if (btn) btn.hidden = true; return; }
    audio = new Audio();
    audio.loop = true;
    audio.preload = 'none';
    var start = +D.music.start_at_sec || 0;
    audio.src = D.music.file + (start > 0 ? '#t=' + start : '');
    audio.addEventListener('error', function () { if (btn) btn.hidden = true; audio = null; });
    if (btn) btn.addEventListener('click', function () { userMuted = !userMuted; if (userMuted) pauseMusic(); else playMusic(); updateMuteBtn(); });
    doc.addEventListener('visibilitychange', function () {
      if (!audio || !opened) return;
      if (doc.hidden) audio.pause(); else if (!userMuted) playMusic();
    });
  }
  function playMusic() {
    if (!audio) return;
    var target = D.music.volume != null ? D.music.volume : 0.8;
    clearInterval(fadeTimer);
    try { audio.volume = 0; } catch (e) { /* iOS: volume is read-only */ }
    var p = audio.play();
    if (p && p.catch) p.catch(function () { userMuted = true; updateMuteBtn(); });
    var v = 0;
    fadeTimer = setInterval(function () {
      v = Math.min(target, v + target / 20);
      try { audio.volume = v; } catch (e) { /* ignore */ }
      if (v >= target) clearInterval(fadeTimer);
    }, 100);
  }
  function pauseMusic() { if (audio) { clearInterval(fadeTimer); audio.pause(); } }
  function updateMuteBtn() {
    var btn = $('#fab-mute');
    if (!btn) return;
    btn.innerHTML = userMuted ? ICON.soundOff : ICON.soundOn;
    btn.setAttribute('aria-label', ui(userMuted ? 'music_on' : 'music_off'));
    btn.setAttribute('aria-pressed', String(!userMuted));
  }

  /* ---------------- Falling petals ---------------- */
  var petals = { canvas: null, ctx: null, list: [], raf: 0, active: false, last: 0, sprites: null, w: 0, h: 0, dpr: 1 };
  function defaultSprites(dpr) {
    var c = doc.createElement('canvas'), s = Math.round(18 * dpr);
    c.width = c.height = s;
    var x = c.getContext('2d');
    x.fillStyle = '#fffdf3';
    x.beginPath();
    x.ellipse(s / 2, s / 2, s * 0.22, s * 0.45, 0, 0, Math.PI * 2);
    x.fill();
    return [c];
  }
  function startPetals() {
    if (reduceMotion || T.petals === false || !window.requestAnimationFrame) return;
    var P = petals;
    P.canvas = h('canvas', { class: 'petals', 'aria-hidden': 'true' });
    body.appendChild(P.canvas);
    P.ctx = P.canvas.getContext('2d');
    P.dpr = Math.min(window.devicePixelRatio || 1, 2);
    P.sprites = (O.petalSprites && O.petalSprites(P.dpr)) || defaultSprites(P.dpr);
    function resize() {
      P.w = window.innerWidth; P.h = window.innerHeight;
      P.canvas.width = Math.round(P.w * P.dpr); P.canvas.height = Math.round(P.h * P.dpr);
    }
    resize();
    window.addEventListener('resize', resize);
    P.active = true;
    // Petals fall while the names scene is on screen, and stop when it scrolls away.
    var hero = $('[data-scene="names"]');
    if (hero && window.IntersectionObserver) {
      new IntersectionObserver(function (en) {
        P.active = en[0].isIntersecting;
        if (P.active) loopPetals();
      }).observe(hero);
    }
    doc.addEventListener('visibilitychange', function () {
      if (doc.hidden) { cancelAnimationFrame(P.raf); P.raf = 0; } else loopPetals();
    });
    loopPetals();
  }
  function spawnPetal(top) {
    var P = petals;
    return {
      x: Math.random() * P.w,
      y: top ? -30 - Math.random() * P.h * 0.6 : -20,
      vy: 38 + Math.random() * 42,
      sway: 14 + Math.random() * 26,
      ph: Math.random() * 6.28,
      rot: Math.random() * 6.28,
      vr: (Math.random() - 0.5) * 1.6,
      sc: 0.6 + Math.random() * 0.7,
      sp: P.sprites[Math.floor(Math.random() * P.sprites.length)]
    };
  }
  function loopPetals() {
    var P = petals;
    if (P.raf || doc.hidden) return;
    P.last = 0;
    function frame(t) {
      var dt = P.last ? Math.min((t - P.last) / 1000, 0.05) : 0.016;
      P.last = t;
      if (P.active && P.list.length < 25 && Math.random() < 0.18) P.list.push(spawnPetal(P.list.length < 6));
      var c = P.ctx;
      c.setTransform(1, 0, 0, 1, 0, 0);
      c.clearRect(0, 0, P.canvas.width, P.canvas.height);
      for (var i = P.list.length - 1; i >= 0; i--) {
        var p = P.list[i];
        p.y += p.vy * dt;
        p.ph += dt * 1.4;
        p.rot += p.vr * dt;
        if (p.y > P.h + 30) { P.list.splice(i, 1); continue; }
        var x = p.x + Math.sin(p.ph) * p.sway;
        var s = p.sp.width * p.sc;
        c.setTransform(P.dpr, 0, 0, P.dpr, 0, 0);
        c.translate(x, p.y);
        c.rotate(p.rot);
        c.drawImage(p.sp, -s / (2 * P.dpr), -s / (2 * P.dpr), s / P.dpr, s / P.dpr);
      }
      if (P.list.length || P.active) P.raf = requestAnimationFrame(frame);
      else { P.raf = 0; c.setTransform(1, 0, 0, 1, 0, 0); c.clearRect(0, 0, P.canvas.width, P.canvas.height); }
    }
    P.raf = requestAnimationFrame(frame);
  }

  /* ---------------- Scroll reveals ---------------- */
  var io = null, rendered = false;
  function observeReveals(ctx) {
    var els = $$('.reveal:not(.in), .polaroid:not(.in), .shimmer:not(.in)', ctx);
    if (!opened) return;
    // A language switch rebuilds cards; show them straight away instead of animating again.
    if (rendered && ctx) { els.forEach(function (e) { e.classList.add('in'); }); return; }
    if (!io) {
      if (!window.IntersectionObserver || reduceMotion) { els.forEach(function (e) { e.classList.add('in'); }); return; }
      io = new IntersectionObserver(function (entries) {
        entries.forEach(function (en) {
          if (en.isIntersecting) { en.target.classList.add('in'); io.unobserve(en.target); }
        });
      }, { rootMargin: '0px 0px -8% 0px', threshold: 0.12 });
    }
    els.forEach(function (e) { io.observe(e); });
  }

  /* ---------------- Watermark ---------------- */
  function applyWatermark() {
    if (!D.watermark) return;
    var brand = String(B.name || '').replace(/[&<>"']/g, '');
    var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="300" height="120"><text x="8" y="66" font-family="Arial,sans-serif" font-size="19" font-weight="700" fill="rgba(255,255,255,.42)" stroke="rgba(0,0,0,.28)" stroke-width=".7">PREVIEW — ' + brand + '</text></svg>';
    var wm = h('div', { class: 'watermark', 'aria-hidden': 'true' });
    wm.style.setProperty('--wm-img', 'url("data:image/svg+xml,' + encodeURIComponent(svg) + '")');
    body.appendChild(wm);
    body.appendChild(h('div', { class: 'preview-ribbon', 'data-ui': 'preview_ribbon' }));
  }

  /* ---------------- Opener ---------------- */
  function openInvite() {
    if (opened) return;
    opened = true;
    var opener = $('#opener');
    if (!userMuted) playMusic();
    updateMuteBtn();
    body.classList.add('is-open');
    if (opener) opener.classList.add('is-opening');
    startPetals();
    if (O.onOpen) safe(O.onOpen);
    var ms = reduceMotion ? 30 : (T.opener_ms || 1400);
    setTimeout(function () { observeReveals(); }, reduceMotion ? 0 : ms * 0.55);
    setTimeout(function () {
      if (opener) { opener.classList.add('is-gone'); opener.setAttribute('aria-hidden', 'true'); }
      body.classList.remove('is-locked');
      var h1 = $('h1');
      if (h1) { h1.setAttribute('tabindex', '-1'); h1.focus({ preventScroll: true }); }
    }, ms);
  }

  /* ---------------- Render all language-dependent parts ---------------- */
  function renderAll() {
    safe(bindText);
    safe(renderLangToggle);
    safe(renderPhotos);
    safe(renderEvents);
    safe(renderCountdown);
    safe(renderRsvpText);
    safe(renderContacts);
    safe(renderFooter);
    safe(updateMuteBtn);
    if (O.onLang) safe(function () { O.onLang(lang); });
    rendered = true;
  }

  /* ---------------- Boot ---------------- */
  function boot() {
    var opener = $('#opener');
    if (opener) {
      opener.addEventListener('click', openInvite);
      opener.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openInvite(); }
      });
    }
    var stored = lsGet(LANG_KEY);
    lang = D.langs.indexOf(stored) >= 0 ? stored : (D.langs.indexOf(D.default_lang) >= 0 ? D.default_lang : D.langs[0]);
    root.setAttribute('lang', lang);
    guest = readGuest();
    safe(function () { $$('.tap-icon').forEach(function (e) { e.innerHTML = ICON.tap; }); });
    safe(function () { var s = $('#fab-share'); if (s) { s.innerHTML = ICON.share; s.addEventListener('click', doShare); } });
    safe(function () { $$('[data-share]').forEach(function (b) { b.addEventListener('click', doShare); }); });
    safe(function () { if (O.init) O.init({ tr: tr, ui: ui, lang: function () { return lang; }, h: h, reduceMotion: reduceMotion }); });
    safe(setupAudio);
    safe(buildRsvp);
    safe(applyWatermark);
    renderAll();
    safe(loadWishes);
    if (!opener) openInvite();
  }

  window.InviteEngine = { tr: tr, ui: ui, setLang: setLang, open: openInvite, sanitizeGuest: readGuest };
  boot();
})();
