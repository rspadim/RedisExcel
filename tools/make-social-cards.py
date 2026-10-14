"""Generates the social card in every site language from the pt-BR original.

Approach: reuse the original image as the artwork (gradient, logo, wordmark and
the URL line are identical in all languages) and replace ONLY the subtitle line:

  1. inpaint the subtitle band column by column, sampling the background a few
     pixels above it - the gradient is smooth, so a per-column fill is seamless
     without having to model the gradient itself;
  2. draw the translated subtitle in Segoe UI at the size calibrated against the
     ORIGINAL width, left-aligned and vertically centred on the original band.

Run from the repository root:  python tools/make-social-cards.py
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

REPO = Path(__file__).resolve().parent.parent
ASSETS = REPO / "docs" / "assets"
ORIGINAL = ASSETS / "social-card.png"

# Segoe UI matches the original card; the fallbacks keep the script usable where
# that font does not exist (the artwork itself is reused, only the subtitle is
# redrawn, so a different face is a cosmetic difference).
FONT_CANDIDATES = [
    Path(r"C:\Windows\Fonts\segoeui.ttf"),
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
]

# The pt-BR subtitle as it appears in the original card, and its measured box.
ORIGINAL_TEXT = "Redis dentro do Excel - RTD, UDF e Pub/Sub ao vivo"
SUB_X0, SUB_Y0, SUB_X1, SUB_Y1 = 376, 368, 1109, 391

TRANSLATIONS = {
    "en": "Redis inside Excel - RTD, UDF and Pub/Sub, live",
    "es": "Redis dentro de Excel - RTD, UDF y Pub/Sub en vivo",
}

SAMPLE_OFFSET = 6      # background is read this many pixels above the band
PAD = 4                # inpaint a little beyond the band on every side


def calibrated_size(draw, font_path, reference, target_width, start=24, stop=40):
    """Font size whose rendering of `reference` is closest to `target_width`."""
    best, best_delta = start, None
    for size in range(start, stop + 1):
        font = ImageFont.truetype(str(font_path), size)
        width = draw.textlength(reference, font=font)
        delta = abs(width - target_width)
        if best_delta is None or delta < best_delta:
            best, best_delta = size, delta
    return best


def main():
    font_path = next((path for path in FONT_CANDIDATES if path.exists()), None)
    if font_path is None:
        raise SystemExit("no usable font found; add one to FONT_CANDIDATES")

    original = Image.open(ORIGINAL).convert("RGB")
    width, height = original.size
    probe = ImageDraw.Draw(original)

    target_width = SUB_X1 - SUB_X0 + 1
    size = calibrated_size(probe, font_path, ORIGINAL_TEXT, target_width)
    font = ImageFont.truetype(str(font_path), size)
    print(f"font: {font_path}")
    print(f"subtitle font size: {size} px "
          f"(renders the pt text {probe.textlength(ORIGINAL_TEXT, font=font):.0f} px wide, "
          f"original {target_width} px)")

    for locale, text in TRANSLATIONS.items():
        card = original.copy()
        draw = ImageDraw.Draw(card)

        # 1) inpaint the subtitle band, column by column.
        for x in range(SUB_X0 - PAD, SUB_X1 + PAD + 1):
            if x < 0 or x >= width:
                continue
            background = card.getpixel((x, SUB_Y0 - SAMPLE_OFFSET))
            for y in range(SUB_Y0 - PAD, SUB_Y1 + PAD + 1):
                if 0 <= y < height:
                    card.putpixel((x, y), background)

        # 2) draw the translated subtitle on the same baseline, left aligned.
        bbox = draw.textbbox((0, 0), text, font=font)
        baseline_y = SUB_Y1 - bbox[3]
        draw.text((SUB_X0, baseline_y), text, font=font, fill=(255, 255, 255))

        out = ASSETS / f"social-card-{locale}.png"
        card.save(out, "PNG", optimize=True)
        print(f"wrote {out.name} ({out.stat().st_size // 1024} KB, "
              f"text {draw.textlength(text, font=font):.0f} px wide)")

    # Keep the old landing-page URL alive with the prettier artwork: the site used
    # to point its hero at redisexcel.jpg (a busy diagram) and that URL may have
    # been shared already. The repository Readme uses the English card directly
    # (docs/assets/social-card-en.png), so no root-level copy is written.
    legacy = ASSETS / "redisexcel.jpg"
    original.save(legacy, "JPEG", quality=92, optimize=True)
    print(f"wrote {legacy.name} (same artwork as the pt card)")


if __name__ == "__main__":
    main()
