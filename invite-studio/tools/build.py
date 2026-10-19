#!/usr/bin/env python3
"""Build dist/ from templates + invite orders.

    python tools/build.py                  # build every invite (and the samples)
    python tools/build.py --only <slug>    # rebuild one invite
    python tools/build.py --watch          # rebuild when a file changes
    python tools/build.py --serve          # build, watch and serve dist/ on your Wi-Fi (port 8080)

Each order is invites/<slug>/invite.json (samples/*.json are built too unless --no-samples).
Data is inlined into the page as window.INVITE, so the page needs no extra request.
Only the standard library is needed for this step.
"""
from __future__ import annotations

import argparse
import gzip
import html
import json
import re
import shutil
import socket
import sys
import threading
import time
from datetime import datetime
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ROOT / "templates"
ENGINE = TEMPLATES / "_engine"
INVITES = ROOT / "invites"
SAMPLES = ROOT / "samples"
DIST = ROOT / "dist"
BRAND_FILE = ROOT / "config" / "brand.json"
SCHEMA_FILE = ENGINE / "invite.schema.json"

# Customer-facing output must never contain the owner-side staff details.
PRIVATE_MARKERS = ["ram lakhan", "lakhan meena", "inspector", "cbic", "gst department"]

BUDGET_CODE_GZ = 400 * 1024     # HTML + CSS + JS + fonts, gzipped
BUDGET_FONTS_EST = 285 * 1024   # worst case, measured 19-10-2026: Tamil first load 114 KB, +Devanagari 127 KB, +Great Vibes 43 KB
BUDGET_TOTAL = 2 * 1024 * 1024  # everything incl. photos
BUDGET_PHOTO = 200 * 1024
BUDGET_OG = 250 * 1024


class BuildError(Exception):
    pass


# ---------------------------------------------------------------- schema validation
def validate(value, schema, root, path="$"):
    """Small JSON-Schema (draft-07 subset) validator so the build needs no extra packages."""
    errors: list[str] = []
    if "$ref" in schema:
        ref = schema["$ref"]
        node = root
        for part in ref.lstrip("#/").split("/"):
            node = node[part]
        return validate(value, node, root, path)
    if "oneOf" in schema:
        passing = [s for s in schema["oneOf"] if not validate(value, s, root, path)]
        if len(passing) != 1:
            errors.append(f"{path}: does not match exactly one allowed shape")
        return errors

    types = schema.get("type")
    if types:
        types = types if isinstance(types, list) else [types]
        if not any(_is_type(value, t) for t in types):
            return [f"{path}: expected {' or '.join(types)}, got {type(value).__name__}"]
    if value is None:
        return errors
    if "enum" in schema and value not in schema["enum"]:
        errors.append(f"{path}: {value!r} is not one of {schema['enum']}")
    if isinstance(value, str):
        if "pattern" in schema and not re.search(schema["pattern"], value):
            errors.append(f"{path}: {value!r} does not match {schema['pattern']}")
        if "maxLength" in schema and len(value) > schema["maxLength"]:
            errors.append(f"{path}: longer than {schema['maxLength']} characters")
        if "minLength" in schema and len(value) < schema["minLength"]:
            errors.append(f"{path}: shorter than {schema['minLength']} characters")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if "minimum" in schema and value < schema["minimum"]:
            errors.append(f"{path}: below minimum {schema['minimum']}")
        if "maximum" in schema and value > schema["maximum"]:
            errors.append(f"{path}: above maximum {schema['maximum']}")
    if isinstance(value, list):
        if "minItems" in schema and len(value) < schema["minItems"]:
            errors.append(f"{path}: needs at least {schema['minItems']} item(s)")
        if "maxItems" in schema and len(value) > schema["maxItems"]:
            errors.append(f"{path}: allows at most {schema['maxItems']} item(s)")
        if schema.get("uniqueItems") and len({json.dumps(v, sort_keys=True) for v in value}) != len(value):
            errors.append(f"{path}: items must be unique")
        if "items" in schema:
            for i, item in enumerate(value):
                errors += validate(item, schema["items"], root, f"{path}[{i}]")
    if isinstance(value, dict):
        for key in schema.get("required", []):
            if key not in value:
                errors.append(f"{path}: missing required field '{key}'")
        if "minProperties" in schema and len(value) < schema["minProperties"]:
            errors.append(f"{path}: must not be empty")
        props = schema.get("properties", {})
        for key, sub in value.items():
            if key in props:
                errors += validate(sub, props[key], root, f"{path}.{key}")
            elif schema.get("additionalProperties") is False:
                errors.append(f"{path}: unknown field '{key}'")
    return errors


def _is_type(value, t):
    return {
        "string": isinstance(value, str),
        "object": isinstance(value, dict),
        "array": isinstance(value, list),
        "boolean": isinstance(value, bool),
        "integer": isinstance(value, int) and not isinstance(value, bool),
        "number": isinstance(value, (int, float)) and not isinstance(value, bool),
        "null": value is None,
    }[t]


def iter_i18n(value, path="$"):
    """Yield (path, obj) for every {ta,en,hi} text object in the invite."""
    if isinstance(value, dict):
        keys = set(value)
        if keys and keys <= {"ta", "en", "hi"} and all(isinstance(v, str) for v in value.values()):
            yield path, value
            return
        for k, v in value.items():
            yield from iter_i18n(v, f"{path}.{k}")
    elif isinstance(value, list):
        for i, v in enumerate(value):
            yield from iter_i18n(v, f"{path}[{i}]")


def check_invite(inv: dict, src_dir: Path, is_sample: bool, schema: dict) -> list[str]:
    """Schema + business-rule checks. Raises BuildError on errors, returns warnings."""
    errors = validate(inv, schema, schema)
    warnings: list[str] = []
    if errors:
        raise BuildError("\n  ".join(["invite.json is not valid:"] + errors[:25]))

    if not is_sample and src_dir.name != inv["slug"]:
        errors.append(f"folder name '{src_dir.name}' must equal slug '{inv['slug']}'")
    tdir = TEMPLATES / inv["template"]
    if not (tdir / "template.json").exists():
        errors.append(f"template '{inv['template']}' not found in templates/")
    else:
        tmeta = json.loads((tdir / "template.json").read_text(encoding="utf-8"))
        variant = inv.get("variant", tmeta.get("default_variant"))
        if variant not in tmeta.get("variants", []):
            errors.append(f"variant '{variant}' not in {tmeta.get('variants')}")
        if inv["occasion"] not in tmeta.get("occasions", []):
            warnings.append(f"occasion '{inv['occasion']}' is not listed for template {tmeta['code']}")
    if inv["default_lang"] not in inv["langs"]:
        errors.append("default_lang must be one of langs")

    ids = [e["id"] for e in inv["events"]]
    if len(ids) != len(set(ids)):
        errors.append("event ids must be unique")
    if inv.get("countdown_to") and inv["countdown_to"] not in ids:
        errors.append(f"countdown_to '{inv['countdown_to']}' is not an event id {ids}")
    for e in inv["events"]:
        if e.get("end") and parse_dt(e["end"]) <= parse_dt(e["start"]):
            errors.append(f"event '{e['id']}': end is not after start")

    for p in inv.get("photos", []):
        src = p if isinstance(p, str) else p["src"]
        if not (src_dir / src).is_file():
            errors.append(f"photo not found: {src_dir / src}")
    music = inv.get("music")
    if music and not resolve_asset(music["file"], src_dir):
        warnings.append(f"music file not found: {music['file']} (music button will hide)")
    rsvp = inv.get("rsvp") or {}
    if rsvp.get("enabled") and not inv.get("watermark", True):
        ep = rsvp.get("endpoint", "")
        if not ep.startswith("https://") or "XXXX" in ep:
            warnings.append("RSVP is enabled but rsvp.endpoint is still a placeholder")
    for c in inv.get("contacts", []):
        for k in ("whatsapp", "phone"):
            if "X" in c.get(k, "").upper():
                warnings.append(f"contact '{c['label'].get('en', '')}' has a placeholder {k} number (button hidden)")

    missing = []
    for path, obj in iter_i18n(inv):
        gaps = [l for l in inv["langs"] if not obj.get(l)]
        if gaps:
            missing.append(f"{path} ({', '.join(gaps)})")
    if missing:
        warnings.append(f"{len(missing)} text(s) missing a language, will fall back to en/ta: " + "; ".join(missing[:4]) + (" …" if len(missing) > 4 else ""))
    if errors:
        raise BuildError("\n  ".join(["invite.json has problems:"] + errors))
    return warnings


def parse_dt(s: str) -> datetime:
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


def resolve_asset(rel: str, src_dir: Path) -> Path | None:
    for base in (ROOT, src_dir):
        p = (base / rel).resolve()
        if p.is_file():
            return p
    return None


# ---------------------------------------------------------------- page assembly
def inline_json(obj) -> str:
    s = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    return (s.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
             .replace("\u2028", "\\u2028").replace("\u2029", "\\u2029"))


def inline_js(path: Path) -> str:
    js = path.read_text(encoding="utf-8")
    if "</script" in js.lower():
        raise BuildError(f"{path} contains '</script' and cannot be inlined")
    return js


def first_text(obj, *langs):
    if not obj:
        return ""
    for l in langs + ("en", "ta", "hi"):
        if obj.get(l):
            return obj[l]
    return ""


def head_html(inv: dict, brand: dict, tmeta: dict) -> str:
    e = partial(html.escape, quote=True)
    domain = brand["domain"].strip("/")
    url = f"https://{domain}/{inv['slug']}/"
    names = first_text(inv["hosts"]["primary"], "en")
    if inv["hosts"].get("secondary"):
        names += " & " + first_text(inv["hosts"]["secondary"], "en")
    og = inv.get("og") or {}
    title = og.get("title") or names
    desc = og.get("description") or "Tap to open the invitation"
    fonts = tmeta.get("fonts_url", "")
    icon = ("data:image/svg+xml," + "%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E"
            "%3Ccircle cx='16' cy='16' r='15' fill='%235B0E14' stroke='%23C9A227' stroke-width='2'/%3E"
            "%3Cpath d='M16 7c3 4 3 7 0 10c-3-3-3-6 0-10z' fill='%23E8D08A'/%3E"
            "%3Crect x='9' y='19' width='14' height='3' rx='1.5' fill='%23C9A227'/%3E%3C/svg%3E")
    lines = [
        f"<title>{e(title)}</title>",
        f'<meta name="description" content="{e(desc)}">',
        '<meta name="robots" content="noindex, nofollow">',
        '<meta name="theme-color" content="#3A070B">',
        '<meta name="format-detection" content="telephone=no">',
        f'<link rel="icon" href="{icon}">',
        f'<meta property="og:title" content="{e(title)}">',
        f'<meta property="og:description" content="{e(desc)}">',
        f'<meta property="og:image" content="{e(url)}og.jpg">',
        '<meta property="og:image:type" content="image/jpeg">',
        '<meta property="og:image:width" content="1200">',
        '<meta property="og:image:height" content="630">',
        f'<meta property="og:image:alt" content="{e(title)}">',
        f'<meta property="og:url" content="{e(url)}">',
        '<meta property="og:type" content="website">',
        f'<meta property="og:site_name" content="{e(brand["name"])}">',
        '<meta name="twitter:card" content="summary_large_image">',
        f'<meta name="twitter:title" content="{e(title)}">',
        f'<meta name="twitter:description" content="{e(desc)}">',
        f'<meta name="twitter:image" content="{e(url)}og.jpg">',
    ]
    if fonts:
        lines += [
            '<link rel="preconnect" href="https://fonts.googleapis.com">',
            '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>',
            # Non-blocking font CSS: text paints at once in a fallback font, then swaps.
            f'<link rel="preload" as="style" href="{e(fonts)}" onload="this.onload=null;this.rel=\'stylesheet\'">',
            f'<noscript><link rel="stylesheet" href="{e(fonts)}"></noscript>',
        ]
    return "\n".join(lines)


def public_brand(brand: dict) -> dict:
    """Only what the page needs; prices and internal settings stay out of customer pages."""
    return {k: brand[k] for k in ("name", "name_i18n", "whatsapp", "domain", "footer") if k in brand}


def build_one(src: Path, brand: dict, schema: dict, is_sample: bool, watermark_override: bool | None) -> dict:
    src_dir = src.parent
    inv = json.loads(src.read_text(encoding="utf-8"))
    warnings = check_invite(inv, src_dir, is_sample, schema)
    if watermark_override is not None:
        inv["watermark"] = watermark_override
    inv.setdefault("watermark", True)

    tdir = TEMPLATES / inv["template"]
    tmeta = json.loads((tdir / "template.json").read_text(encoding="utf-8"))
    variant = inv.get("variant") or tmeta["default_variant"]
    slug = inv["slug"]
    out = DIST / slug
    keep = {"og.jpg", "status.jpg"}  # generated images survive a page rebuild
    if out.exists():
        for child in out.iterdir():
            if child.name in keep:
                continue
            shutil.rmtree(child) if child.is_dir() else child.unlink()
    out.mkdir(parents=True, exist_ok=True)

    # Copy photos (Phase 2 replaces this with compression to WebP).
    page_data = json.loads(json.dumps(inv))
    photo_bytes = 0
    if inv.get("photos"):
        (out / "photos").mkdir(exist_ok=True)
        seen = set()
        new = []
        for p in inv["photos"]:
            obj = {"src": p} if isinstance(p, str) else dict(p)
            name = Path(obj["src"]).name
            if name in seen:
                raise BuildError(f"two photos share the file name {name}")
            seen.add(name)
            shutil.copy2(src_dir / obj["src"], out / "photos" / name)
            size = (out / "photos" / name).stat().st_size
            photo_bytes += size
            if size > BUDGET_PHOTO:
                warnings.append(f"photo {name} is {kb(size)} (> {kb(BUDGET_PHOTO)}); Phase 2 compression will fix this")
            obj["src"] = f"photos/{name}"
            new.append(obj)
        page_data["photos"] = new

    music_bytes = 0
    music = inv.get("music")
    if music:
        mpath = resolve_asset(music["file"], src_dir)
        if mpath:
            music_bytes = mpath.stat().st_size
            try:
                rel = mpath.relative_to(ROOT / "assets")
                dest = DIST / "assets" / rel
                page_data["music"]["file"] = "../assets/" + rel.as_posix()
            except ValueError:
                dest = out / "music" / mpath.name
                page_data["music"]["file"] = "music/" + mpath.name
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(mpath, dest)
        else:
            page_data["music"] = None

    page = (tdir / "index.html").read_text(encoding="utf-8")
    scripts = "\n".join([
        "<script>window.INVITE=" + inline_json(page_data) + ";",
        "window.BRAND=" + inline_json(public_brand(brand)) + ";",
        "window.TEMPLATE=" + inline_json({k: tmeta[k] for k in ("code", "opener_ms", "petals") if k in tmeta}) + ";</script>",
        "<script>" + inline_js(tdir / "ornaments.js") + "</script>",
        "<script>" + inline_js(ENGINE / "engine.js") + "</script>",
    ])
    replacements = {
        "{{LANG}}": inv["default_lang"],
        "{{VARIANT}}": variant,
        "<!--@HEAD-->": head_html(inv, brand, tmeta),
        "/*@ENGINE_CSS*/": (ENGINE / "engine.css").read_text(encoding="utf-8"),
        "/*@THEME_CSS*/": (tdir / "theme.css").read_text(encoding="utf-8"),
        "<!--@SCRIPTS-->": scripts,
    }
    for k, v in replacements.items():
        if k not in page:
            raise BuildError(f"template {inv['template']}/index.html is missing the {k} slot")
        page = page.replace(k, v)

    lowered = page.lower()
    leaks = [m for m in PRIVATE_MARKERS if m in lowered]
    if leaks:
        raise BuildError(f"customer page would contain private details: {leaks}. Remove them from the data/template.")

    (out / "index.html").write_text(page, encoding="utf-8")
    raw = len(page.encode("utf-8"))
    gz = len(gzip.compress(page.encode("utf-8"), 9))
    og = out / "og.jpg"
    if not og.exists():
        warnings.append("og.jpg not generated yet (Phase 2: tools/images.py) — WhatsApp preview will have no image")
    elif og.stat().st_size > BUDGET_OG:
        warnings.append(f"og.jpg is {kb(og.stat().st_size)} (> 250 KB, WhatsApp may drop it)")
    code_total = gz + BUDGET_FONTS_EST
    total = gz + BUDGET_FONTS_EST + photo_bytes
    if code_total > BUDGET_CODE_GZ:
        warnings.append(f"code+fonts ≈ {kb(code_total)} is over the {kb(BUDGET_CODE_GZ)} budget")
    if total > BUDGET_TOTAL:
        warnings.append(f"page total ≈ {kb(total)} is over the {kb(BUDGET_TOTAL)} budget")
    return {
        "slug": slug, "template": tmeta["code"], "variant": variant, "watermark": inv["watermark"],
        "html_raw": raw, "html_gz": gz, "photos": photo_bytes, "n_photos": len(inv.get("photos", [])),
        "music": music_bytes, "total": total, "warnings": warnings,
    }


def write_root_page(brand: dict):
    """Tiny brand landing page so the bare domain never lists invite folders."""
    e = partial(html.escape, quote=True)
    wa = re.sub(r"\D", "", brand.get("whatsapp", ""))
    page = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>{e(brand['name'])}</title>
<meta name="robots" content="noindex">
<style>body{{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#3A070B;color:#FFF8E7;font:18px/1.6 Georgia,serif;text-align:center;padding:24px}}a{{color:#E8D08A}}h1{{color:#E8D08A;font-weight:600}}</style>
</head><body><main><h1>{e(brand['name'])}</h1><p>{e(brand['name_i18n'].get('ta', ''))}</p>
<p>Digital invitations for weddings and family functions.</p>
<p><a href="https://wa.me/{wa}">WhatsApp {e(brand.get('whatsapp', ''))}</a></p></main></body></html>"""
    (DIST / "index.html").write_text(page, encoding="utf-8")


# ---------------------------------------------------------------- driver
def kb(n: int) -> str:
    return f"{n / 1024:.0f} KB" if n < 1024 * 1024 else f"{n / 1024 / 1024:.2f} MB"


def sources(include_samples: bool) -> list[tuple[Path, bool]]:
    found = [(p, False) for p in sorted(INVITES.glob("*/invite.json"))]
    if include_samples:
        found += [(p, True) for p in sorted(SAMPLES.glob("*.json"))]
    return found


def build_all(args) -> bool:
    brand = json.loads(BRAND_FILE.read_text(encoding="utf-8"))
    schema = json.loads(SCHEMA_FILE.read_text(encoding="utf-8"))
    wm = None if args.watermark is None else args.watermark == "on"
    DIST.mkdir(exist_ok=True)
    write_root_page(brand)
    ok = True
    built = 0
    print(f"\nBuilding → {DIST.relative_to(ROOT)}/   ({datetime.now():%d-%m-%Y %H:%M:%S})")
    for src, is_sample in sources(not args.no_samples):
        label = src.relative_to(ROOT)
        try:
            slug_hint = json.loads(src.read_text(encoding="utf-8")).get("slug")
        except json.JSONDecodeError as ex:
            print(f"  ✗ {label}: not valid JSON ({ex})")
            ok = False
            continue
        if args.only and slug_hint != args.only:
            continue
        try:
            r = build_one(src, brand, schema, is_sample, wm)
        except BuildError as ex:
            print(f"  ✗ {label}\n  {ex}")
            ok = False
            continue
        built += 1
        print(f"  ✓ {r['slug']}  [{r['template']} · {r['variant']} · watermark {'ON' if r['watermark'] else 'off'}]")
        print(f"      page {kb(r['html_raw'])} raw / {kb(r['html_gz'])} gzip · "
              f"+ fonts ≈ {kb(BUDGET_FONTS_EST)} = {kb(r['html_gz'] + BUDGET_FONTS_EST)} (budget {kb(BUDGET_CODE_GZ)})")
        print(f"      photos {r['n_photos']} = {kb(r['photos'])} · total ≈ {kb(r['total'])} (budget {kb(BUDGET_TOTAL)})"
              f" · music {kb(r['music'])} (loads after tap)")
        for w in r["warnings"]:
            print(f"      ! {w}")
    if args.only and not built:
        print(f"  ✗ no invite with slug '{args.only}'")
        ok = False
    return ok


def snapshot() -> dict:
    files = {}
    for base in (TEMPLATES, INVITES, SAMPLES, ROOT / "config", ROOT / "assets"):
        for p in base.rglob("*"):
            if p.is_file():
                files[p] = p.stat().st_mtime
    return files


def lan_ip() -> str:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))  # no packet is sent; this just picks the Wi-Fi interface
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


def serve(port: int):
    handler = partial(SimpleHTTPRequestHandler, directory=str(DIST))
    handler.log_message = lambda *a, **k: None
    httpd = ThreadingHTTPServer(("0.0.0.0", port), handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    ip = lan_ip()
    print(f"\nServing dist/ on http://{ip}:{port}/  (phone on the same Wi-Fi)")
    for src, _ in sources(True):
        try:
            slug = json.loads(src.read_text(encoding="utf-8"))["slug"]
            print(f"  http://{ip}:{port}/{slug}/      ·  with guest name: http://{ip}:{port}/{slug}/?to=Ravi-Uncle")
        except (json.JSONDecodeError, KeyError):
            pass


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--only", metavar="SLUG", help="build just this invite")
    ap.add_argument("--watch", action="store_true", help="rebuild when templates/invites change")
    ap.add_argument("--serve", nargs="?", const=8080, type=int, metavar="PORT", help="serve dist/ on the LAN (implies --watch)")
    ap.add_argument("--no-samples", action="store_true", help="skip samples/*.json")
    ap.add_argument("--watermark", choices=["on", "off"], help="override the watermark flag for this build")
    args = ap.parse_args()

    ok = build_all(args)
    if args.serve:
        serve(args.serve)
        args.watch = True
    if not args.watch:
        sys.exit(0 if ok else 1)
    print("\nWatching for changes… (Ctrl+C to stop)")
    before = snapshot()
    try:
        while True:
            time.sleep(1)
            now = snapshot()
            if now != before:
                before = now
                build_all(args)
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
