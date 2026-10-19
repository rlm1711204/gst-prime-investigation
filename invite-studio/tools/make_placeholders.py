#!/usr/bin/env python3
"""Generate the placeholder assets used by the sample invite. Everything here is drawn
or synthesised in code, so it is original work (see ASSETS_LICENSES.md).

    python tools/make_placeholders.py

Writes samples/photos/01-04.jpg and assets/music/placeholder-pluck.mp3.
Needs Pillow + numpy, and ffmpeg on PATH for the MP3.
"""
from __future__ import annotations

import math
import random
import shutil
import subprocess
import wave
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent.parent
PHOTOS = ROOT / "samples" / "photos"
MUSIC = ROOT / "assets" / "music"
W, H = 1080, 1350


# ---------------------------------------------------------------- photos
def gradient(top, bottom):
    img = Image.new("RGB", (W, H))
    px = ImageDraw.Draw(img)
    for y in range(H):
        t = y / (H - 1)
        px.line([(0, y), (W, y)], fill=tuple(int(a + (b - a) * t) for a, b in zip(top, bottom)))
    return img


def label(img, text):
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 34)
    except OSError:
        font = ImageFont.load_default()
    w = d.textlength(text, font=font)
    d.rounded_rectangle([(W - w) / 2 - 24, H - 110, (W + w) / 2 + 24, H - 50], 28, fill=(0, 0, 0, 120))
    d.text(((W - w) / 2, H - 100), text, font=font, fill=(255, 248, 231))


def couple(d, cx, base, scale, colour_a, colour_b):
    # two simple figures side by side (abstract silhouettes, not real people)
    for dx, col, hh in ((-70, colour_a, 1.0), (70, colour_b, 1.08)):
        x = cx + dx * scale
        r = 52 * scale
        top = base - 520 * scale * hh
        d.ellipse([x - r, top, x + r, top + 2 * r], fill=col)
        d.polygon([(x - 120 * scale, base), (x - 70 * scale, top + 2.4 * r), (x + 70 * scale, top + 2.4 * r), (x + 120 * scale, base)], fill=col)


def photo1():
    img = gradient((250, 196, 120), (122, 22, 32))
    d = ImageDraw.Draw(img)
    d.ellipse([290, 260, 790, 760], fill=(255, 226, 160))
    # temple arch frame
    d.rectangle([140, 380, 220, H], fill=(70, 12, 18))
    d.rectangle([860, 380, 940, H], fill=(70, 12, 18))
    d.pieslice([140, 120, 940, 760], 180, 360, fill=None, outline=(201, 162, 39), width=26)
    couple(d, 540, 1220, 1.0, (60, 10, 16), (40, 6, 10))
    label(img, "Sample photo 1")
    return img


def photo2():
    img = gradient((46, 125, 50), (10, 46, 28))
    d = ImageDraw.Draw(img)
    rnd = random.Random(2)
    for _ in range(26):
        x, y, r = rnd.randint(60, W - 60), rnd.randint(80, H - 200), rnd.randint(40, 90)
        for k, col in enumerate(((232, 120, 26), (242, 163, 27), (255, 210, 122))):
            rr = r * (1 - k * 0.3)
            d.ellipse([x - rr, y - rr, x + rr, y + rr], fill=col)
    img = img.filter(ImageFilter.GaussianBlur(1.2))
    label(img, "Sample photo 2")
    return img


def photo3():
    img = gradient((255, 214, 150), (184, 92, 56))
    d = ImageDraw.Draw(img)
    # gopuram silhouette: stacked trapezoids
    base_y, width = 1150, 620
    for i in range(7):
        w = width * (1 - i * 0.12)
        h = 110
        y = base_y - i * h
        d.polygon([(540 - w / 2, y), (540 - w / 2 + 30, y - h), (540 + w / 2 - 30, y - h), (540 + w / 2, y)], fill=(92, 30, 24))
    d.rectangle([470, base_y - 7 * 110 - 70, 610, base_y - 7 * 110], fill=(92, 30, 24))
    d.rectangle([0, base_y, W, H], fill=(70, 22, 18))
    label(img, "Sample photo 3")
    return img


def photo4():
    img = gradient((58, 7, 11), (20, 3, 5))
    d = ImageDraw.Draw(img)
    for i, x in enumerate(range(140, W, 200)):
        y = 900 + (i % 2) * 60
        glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        g = ImageDraw.Draw(glow)
        g.ellipse([x - 120, y - 260, x + 120, y - 20], fill=(255, 190, 80, 90))
        img.paste(glow.filter(ImageFilter.GaussianBlur(40)), (0, 0), glow.filter(ImageFilter.GaussianBlur(40)))
        d = ImageDraw.Draw(img)
        d.polygon([(x, y - 190), (x + 26, y - 110), (x, y - 80), (x - 26, y - 110)], fill=(255, 196, 60))
        d.pieslice([x - 80, y - 110, x + 80, y + 10], 0, 180, fill=(201, 162, 39))
    label(img, "Sample photo 4")
    return img


def make_photos():
    PHOTOS.mkdir(parents=True, exist_ok=True)
    for i, fn in enumerate((photo1, photo2, photo3, photo4), 1):
        out = PHOTOS / f"{i:02d}.jpg"
        fn().save(out, "JPEG", quality=82, optimize=True, progressive=True)
        print(f"  {out.relative_to(ROOT)}  {out.stat().st_size // 1024} KB")


# ---------------------------------------------------------------- music
SR = 22050


def pluck(freq, dur, decay=0.996, bright=0.5):
    """Karplus–Strong plucked string."""
    n = int(SR * dur)
    period = max(2, int(SR / freq))
    rng = np.random.default_rng(int(freq * 100))
    buf = rng.uniform(-1, 1, period)
    out = np.empty(n)
    for i in range(n):
        j = i % period
        out[i] = buf[j]
        buf[j] = decay * (bright * buf[j] + (1 - bright) * buf[(j + 1) % period])
    env = np.minimum(1, np.arange(n) / (SR * 0.004))
    return out * env


def make_music():
    MUSIC.mkdir(parents=True, exist_ok=True)
    sa = 196.0  # G3
    ratio = {"S": 1, "R": 9 / 8, "G": 5 / 4, "P": 3 / 2, "D": 5 / 3, "S'": 2, "R'": 9 / 4, "G'": 5 / 2}
    # Phrases in the pentatonic Mohanam scale (S R G P D): (note, beats)
    phrases = [
        [("S", 1), ("R", 1), ("G", 1), ("P", 1), ("D", 2), ("P", 2)],
        [("G", 1), ("P", 1), ("D", 1), ("S'", 1), ("D", 2), ("P", 2)],
        [("G", 1), ("R", 1), ("G", 1), ("P", 1), ("G", 2), ("R", 2)],
        [("S", 1), ("R", 1), ("G", 2), ("R", 1), ("S", 1), ("S", 2)],
        [("P", 1), ("D", 1), ("S'", 1), ("R'", 1), ("G'", 2), ("R'", 2)],
        [("S'", 1), ("D", 1), ("P", 1), ("G", 1), ("R", 2), ("S", 2)],
    ]
    beat = 0.42
    total_beats = sum(b for ph in phrases for _, b in ph) * 2
    length = int(SR * (total_beats * beat + 2))
    mix = np.zeros(length)
    t = 0.0
    for ph in phrases * 2:
        for note, beats in ph:
            s = pluck(sa * ratio[note], beats * beat + 1.2, decay=0.9975, bright=0.55) * 0.55
            i = int(t * SR)
            mix[i:i + len(s)] += s[: max(0, length - i)]
            t += beats * beat
    # soft tanpura-like drone: Pa, Sa', Sa', Sa (one octave down)
    drone = [sa * 0.75, sa, sa, sa * 0.5]
    t, k = 0.0, 0
    while t < length / SR - 2:
        s = pluck(drone[k % 4], 2.6, decay=0.9992, bright=0.7) * 0.22
        i = int(t * SR)
        mix[i:i + len(s)] += s[: max(0, length - i)]
        t += 0.62
        k += 1
    fade = int(SR * 1.5)
    mix[:fade] *= np.linspace(0, 1, fade)
    mix[-fade:] *= np.linspace(1, 0, fade)
    mix = mix / np.max(np.abs(mix)) * 0.85
    wav = MUSIC / "placeholder-pluck.wav"
    with wave.open(str(wav), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(SR)
        w.writeframes((mix * 32767).astype("<i2").tobytes())
    mp3 = MUSIC / "placeholder-pluck.mp3"
    if shutil.which("ffmpeg"):
        subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(wav), "-ac", "1", "-b:a", "64k", str(mp3)], check=True)
        wav.unlink()
        print(f"  {mp3.relative_to(ROOT)}  {mp3.stat().st_size // 1024} KB, {length / SR:.0f} s")
    else:
        print(f"  ffmpeg not found; left {wav.relative_to(ROOT)} (convert it to MP3 before use)")


if __name__ == "__main__":
    print("Placeholder photos:")
    make_photos()
    print("Placeholder music:")
    make_music()
