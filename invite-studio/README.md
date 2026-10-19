# Invite Studio

Mobile-first digital invitations shared on WhatsApp as a link. One customer order is one
`invite.json`; the build turns it into a single fast page in `dist/<slug>/`, hosted on
Cloudflare Pages.

> Keep this repository **Private**. Order files contain customers' names, phone numbers and photos.

## What you need

- Python 3.10+ (nothing else for building pages).
- For regenerating the placeholder photos/music only: `pip install pillow numpy` and ffmpeg.
- Optional automated phone checks: Node.js + Playwright (`npm i -g playwright`).

## Try the sample on your phone (same Wi-Fi)

```bash
python tools/build.py --serve
```

It builds everything, watches for changes and prints links like:

```
Serving dist/ on http://192.168.1.23:8080/
  http://192.168.1.23:8080/ramya-karthik-x7q2/   ·  with guest name: …/?to=Ravi-Uncle
```

Open that link in Chrome on your phone, or send it to yourself on WhatsApp and tap it there.
On Windows, click **Allow** when the firewall asks about Python on *private* networks.
The plain alternative is `python tools/build.py` then `python -m http.server -d dist 8080`
(find your laptop's IP with `ipconfig` on Windows or `ip addr` on Linux).

## Create an order

1. Copy `samples/sample-wedding.json` to `invites/<slug>/invite.json`.
   The slug is lowercase names plus a random 4-character suffix, e.g. `priya-arun-k3m9`
   (the suffix stops people guessing other links). The folder name must equal the slug.
2. Put the photos in `invites/<slug>/photos/` and list them under `photos`.
3. Fill every text as `{ "ta": "…", "en": "…", "hi": "…" }`. A missing language falls back to
   English, then Tamil; the build lists any gaps.
4. `python tools/build.py --only <slug>` and check it on your phone.
5. Send the customer the preview (`"watermark": true`: diagonal PREVIEW marks, RSVP switched off).
6. After approval set `"watermark": false`, rebuild and deploy.

The full format is in `templates/_engine/invite.schema.json`; the build validates against it.

### Useful fields

| Field | Notes |
|---|---|
| `template`, `variant` | `kanjeevaram-gold` with `maroon-gold`, `green-gold` or `ivory-gold` |
| `default_lang`, `langs` | language the page opens in, and the toggle (`ta`, `en`, `hi`) |
| `events[].start/end` | ISO time with offset, e.g. `2026-12-11T06:00:00+05:30`; shown exactly as written |
| `events[].traditional_date` | optional Tamil-calendar date typed by hand, e.g. `கார்த்திகை 25` |
| `countdown_to` | the event `id` to count down to |
| `hosts.joiner` | word between names; default `weds` / `&` / `संग` |
| `rsvp.endpoint` | Google Apps Script URL (Phase 3); placeholder keeps the form switched off |
| `guest_personalisation` | `?to=Ravi-Uncle` shows "Dear Ravi Uncle" (letters and spaces only, max 40) |

## Build commands

```
python tools/build.py                   # build all invites + samples
python tools/build.py --only <slug>     # one invite
python tools/build.py --watch           # rebuild on change
python tools/build.py --serve [PORT]    # build + watch + serve on the LAN
python tools/build.py --watermark off   # override the watermark flag for a test build
python tools/build.py --no-samples      # skip samples/
node tools/qa_check.cjs <slug> [http://localhost:8080] [screenshot-dir]   # automated phone checks
```

The build log prints the page size (raw and gzip), photo total and warnings for missing
languages, placeholder numbers, oversized photos and a missing `og.jpg`. It refuses to write a
page that contains the private staff markers listed in `PRIVATE_MARKERS` in `tools/build.py`.

### Size budget (measured 19-10-2026)

| Part | Size |
|---|---|
| Page (HTML + CSS + JS + data, gzip) | ~26 KB |
| Fonts, Tamil first load (Noto Serif Tamil + Cormorant) | ~114 KB |
| + Hindi selected (Noto Serif Devanagari) | +127 KB |
| + English selected (Great Vibes) | +43 KB |
| Sample photos (4) | 172 KB |
| Music (after the tap only) | 331 KB |

Budget: code + fonts under 400 KB, everything with photos under 2 MB.

## Layout

```
config/brand.json            brand name, business WhatsApp, domain, footer text, prices
templates/_engine/           engine.js, engine.css, invite.schema.json (shared by all templates)
templates/kanjeevaram-gold/  index.html, theme.css, ornaments.js, template.json (KG-01)
invites/<slug>/              invite.json + photos/ for each order
samples/                     sample-wedding.json + generated placeholder photos
assets/music/                royalty-free music (placeholder generated in code)
tools/build.py               the build
tools/make_placeholders.py   regenerates sample photos and placeholder music
tools/qa_check.cjs           Playwright checks at 320/375/412 px
dist/                        build output (not committed)
```

## How the page works

- Full-screen temple-door opener; the tap starts the music (browsers block autoplay), opens the
  doors and starts the falling jasmine. The mute button appears afterwards.
- Scenes: blessing, names with gold shimmer, polaroid photos with a swipeable full-screen viewer,
  event timeline (Maps, Google Calendar, .ics), countdown, RSVP and wishes, family contacts,
  closing, and a footer with the brand's WhatsApp.
- Language toggle switches every string without reloading and remembers the choice.
- `prefers-reduced-motion` turns off petals, door animation and scroll motion.
- Every page has `noindex, nofollow` and full Open Graph tags (the `og.jpg` image comes in Phase 2).
