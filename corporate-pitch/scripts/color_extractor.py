#!/usr/bin/env python3
"""Extract dominant colors from a logo image, excluding black and white."""
import argparse, os, sys

FALLBACK_COLOR = "#1A365D"  # Corporate blue fallback


def extract_colors(logo_path, n_colors=3):
    """Return top N non-black/non-white hex colors from logo image."""
    try:
        from PIL import Image
    except ImportError:
        print(f"[WARN] Pillow not installed. Using fallback: {FALLBACK_COLOR}")
        return [FALLBACK_COLOR]

    if not os.path.exists(logo_path):
        print(f"[WARN] Logo not found: {logo_path}. Using fallback: {FALLBACK_COLOR}")
        return [FALLBACK_COLOR]

    try:
        img = Image.open(logo_path).convert("RGB")
        # Resize for speed
        img = img.resize((100, 100), Image.LANCZOS)

        # Count pixel colors
        pixels = list(img.getdata())
        color_counts = {}
        for r, g, b in pixels:
            # Skip near-black (< 30) and near-white (> 225)
            if r < 30 and g < 30 and b < 30:
                continue
            if r > 225 and g > 225 and b > 225:
                continue
            # Skip very low saturation (grays)
            if abs(r - g) < 15 and abs(g - b) < 15 and abs(r - b) < 15:
                continue
            # Quantize to reduce noise
            qr, qg, qb = (r // 16) * 16, (g // 16) * 16, (b // 16) * 16
            key = (qr, qg, qb)
            color_counts[key] = color_counts.get(key, 0) + 1

        if not color_counts:
            print(f"[WARN] No non-BW colors found. Using fallback: {FALLBACK_COLOR}")
            return [FALLBACK_COLOR]

        # Sort by frequency, take top N
        sorted_colors = sorted(color_counts.items(), key=lambda x: -x[1])
        result = []
        for (r, g, b), _ in sorted_colors[:n_colors]:
            result.append(f"#{r:02X}{g:02X}{b:02X}")

        return result

    except Exception as e:
        print(f"[ERR] Color extraction failed: {e}. Using fallback: {FALLBACK_COLOR}")
        return [FALLBACK_COLOR]


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Extract dominant colors from logo")
    ap.add_argument("--logo", required=True, help="Path to logo image")
    ap.add_argument("--n", type=int, default=3, help="Number of colors to extract")
    args = ap.parse_args()

    colors = extract_colors(args.logo, args.n)
    print(f"[COLORS] {', '.join(colors)}")
    # Output primary (first) color for scripting
    print(colors[0])
