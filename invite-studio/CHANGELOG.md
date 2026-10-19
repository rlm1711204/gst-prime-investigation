# Changelog

## v0.1 — 19-10-2026

Phase 1: skeleton, engine and template KG-01.

- Folder structure, `config/brand.json` (placeholder brand "Vizha Invites", placeholder number).
- `templates/_engine/invite.schema.json`: order format; every text is `{ta, en, hi}`.
- `templates/_engine/engine.js` + `engine.css`: opener, music after tap with fade-in and mute
  button, falling petals (max 25, paused when the tab is hidden or the names scroll away),
  scroll reveals, names shimmer, polaroid photos + swipe viewer, event timeline with Maps,
  Google Calendar and .ics, countdown (switches to "Thank you for blessing us" afterwards),
  RSVP form + wishes list, contacts (call / WhatsApp), footer, share button, language toggle
  remembered in localStorage, `?to=` guest name (sanitised), watermark mode that disables RSVP,
  reduced-motion support.
- Tamil/Hindi month and weekday names and time-of-day words (காலை / மாலை, सुबह / शाम) written in code.
- Template `kanjeevaram-gold` (KG-01) with `maroon-gold`, `green-gold` and `ivory-gold` variants:
  carved temple doors, mango-leaf thoranam, kuthuvilakku lamps, kolam corners, zari border
  dividers, silk texture from CSS gradients, jasmine/marigold petals; all drawn in code.
- `tools/build.py`: schema + business-rule validation, inline data, OG/Twitter tags, size report,
  privacy guard, `--only`, `--watch`, `--serve`, `--watermark`, `--no-samples`.
- `tools/make_placeholders.py`: generated sample photos and a synthesised placeholder tune.
- `tools/qa_check.cjs`: Playwright checks at 320/375/412 px (30 checks).
- Font loading tuned: toggle labels use the phone's fonts and Latin text comes from Cormorant, so
  only the fonts actually used are downloaded (Tamil first load 327 KB → 114 KB).
- `samples/sample-wedding.json`: Ramya & Karthik, three events in Tirunelveli, all three languages.
