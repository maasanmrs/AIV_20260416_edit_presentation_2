#!/usr/bin/env python3
"""
Generate images using Nano Banana Pro (Gemini 3 Pro Image Preview).
Model: gemini-3-pro-image-preview

Features:
- Exact pixel dimensions (no cropping needed)
- Japanese text output with Noto Sans JP font specification
- Business diagram / consulting-style infographic generation
- Style-aware prompt engineering
"""
import argparse, os, sys, time

API_KEY = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
if not API_KEY:
    print("[WARN] GEMINI_API_KEY or GOOGLE_API_KEY not set. Image generation disabled.")

# ── Style suffixes ─────────────────────────────────────────────────────
# Each style adds domain-specific visual instructions to the base prompt.

STYLE_SUFFIXES = {
    "corporate": (
        ", formal corporate photography, clean minimal composition, "
        "professional business environment, high resolution, subtle and elegant, "
        "modern office aesthetic"
    ),
    "technology": (
        ", futuristic technology visualization, clean digital interface, "
        "neon accents on dark background, high resolution, sophisticated and minimal"
    ),
    "infrastructure": (
        ", industrial infrastructure photography, engineering precision, "
        "dramatic scale, clean composition, professional documentary style"
    ),
    "abstract": (
        ", abstract geometric patterns, professional color palette, "
        "clean modern design, subtle gradient, high resolution, corporate aesthetic"
    ),
    "diagram": (
        ", professional business consulting slide diagram, clean vector-style "
        "infographic, flat design icons, white background, "
        "McKinsey/BCG style business concept visualization, "
        "labeled boxes and arrows, professional color scheme with red and black accents, "
        "high resolution"
    ),
}

# ── Japanese business diagram prompt template ──────────────────────────
# Used when style="diagram" to produce consulting-grade Japanese infographics.

JAPANESE_DIAGRAM_PREFIX = (
    "Create a professional business consulting diagram for a presentation slide. "
    "ALL text labels, titles, and annotations MUST be written in Japanese (日本語). "
    "Use the font 'Noto Sans JP' or a clean sans-serif Japanese font for all text. "
    "The diagram should look like a slide from McKinsey, BCG, or Deloitte — "
    "clean flat-design icons, labeled boxes with Japanese text, directional arrows, "
    "professional color scheme (black, white, red accents). "
    "Do NOT include any English text. "
)

JAPANESE_TEXT_INSTRUCTION = (
    " ALL text, labels, captions, and annotations in the image MUST be in Japanese (日本語). "
    "Use 'Noto Sans JP' or a clean Japanese sans-serif font. "
    "Do NOT include any English text in the image. "
)

MODEL_ID = "gemini-3-pro-image-preview"


def generate_image(prompt, output_path, width_px=1280, height_px=720,
                   style="corporate", japanese_text=True):
    """Generate a single image at exact pixel dimensions.

    Args:
        prompt: Image description (will be enhanced with style suffix)
        output_path: Where to save the JPEG
        width_px: Exact output width in pixels
        height_px: Exact output height in pixels
        style: One of corporate/technology/infrastructure/abstract/diagram
        japanese_text: If True, instruct model to use Japanese text + Noto Sans JP

    Returns:
        output_path on success, None on failure.
    """
    if not API_KEY:
        print(f"  [SKIP] No API key — {os.path.basename(output_path)}")
        return None

    try:
        from google import genai
        from google.genai import types
    except ImportError:
        print("[ERR] Run: pip install google-genai")
        return None

    # ── Build the full prompt ──────────────────────────────────────────
    # 1. For diagram style, prepend the Japanese diagram template
    # 2. Add style suffix
    # 3. Add exact dimension instruction
    # 4. Add Japanese text instruction if enabled

    if style == "diagram":
        full_prompt = JAPANESE_DIAGRAM_PREFIX + prompt.rstrip(".")
    else:
        full_prompt = prompt.rstrip(".")

    suffix = STYLE_SUFFIXES.get(style, STYLE_SUFFIXES["corporate"])
    full_prompt += suffix

    # Exact pixel dimensions — prevents need for cropping/trimming
    full_prompt += f", exact output resolution {width_px}x{height_px} pixels"

    # Japanese text instruction for non-diagram styles (diagram already has it)
    if japanese_text and style != "diagram":
        full_prompt += JAPANESE_TEXT_INSTRUCTION

    client = genai.Client(api_key=API_KEY)

    print(f"[IMG] {os.path.basename(output_path)} ({width_px}x{height_px}px, {style})")

    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=full_prompt,
                config=types.GenerateContentConfig(
                    response_modalities=["IMAGE", "TEXT"],
                ),
            )
            # Find image part in response
            image_data = None
            for part in response.candidates[0].content.parts:
                if part.inline_data and part.inline_data.data:
                    image_data = part.inline_data.data
                    break

            if image_data is None:
                print(f"  [WARN] No image in response, retry {attempt + 1}/3...")
                time.sleep(2)
                continue

            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

            # Save via Pillow — resize to exact dimensions if model output differs
            try:
                from PIL import Image
                import io, base64

                if isinstance(image_data, str):
                    raw = base64.b64decode(image_data)
                else:
                    raw = image_data
                img = Image.open(io.BytesIO(raw)).convert("RGB")

                # Resize to exact target if the model didn't match perfectly
                if img.size != (width_px, height_px):
                    img = img.resize((width_px, height_px), Image.LANCZOS)

                img.save(output_path, "JPEG", quality=92)
            except ImportError:
                import base64

                with open(output_path, "wb") as f:
                    if isinstance(image_data, str):
                        f.write(base64.b64decode(image_data))
                    else:
                        f.write(image_data)

            size_kb = os.path.getsize(output_path) // 1024
            print(f"  [OK] {size_kb}KB ({width_px}x{height_px})")
            return output_path

        except Exception as e:
            print(f"  [ERR] attempt {attempt + 1}/3: {e}")
            if attempt < 2:
                time.sleep((attempt + 1) * 3)

    print(f"  [FAIL] Could not generate: {os.path.basename(output_path)}")
    return None


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Generate image via Nano Banana Pro")
    p.add_argument("--prompt", required=True, help="Image generation prompt")
    p.add_argument("--output", required=True, help="Output image path")
    p.add_argument("--width", type=int, default=1280, help="Exact width in pixels")
    p.add_argument("--height", type=int, default=720, help="Exact height in pixels")
    p.add_argument("--aspect-ratio", default=None,
                   help="Aspect ratio (e.g. 16:9). Ignored if --width/--height given explicitly.")
    p.add_argument(
        "--style",
        default="corporate",
        choices=["corporate", "technology", "infrastructure", "abstract", "diagram"],
        help="Image style",
    )
    p.add_argument("--no-japanese", action="store_true",
                   help="Disable Japanese text instruction (default: Japanese enabled)")
    args = p.parse_args()

    # If aspect-ratio is given but width/height are defaults, compute dimensions
    if args.aspect_ratio and args.width == 1280 and args.height == 720:
        parts = args.aspect_ratio.split(":")
        if len(parts) == 2:
            aw, ah = int(parts[0]), int(parts[1])
            # Keep height at 720, compute width
            args.width = round(720 * aw / ah)

    result = generate_image(
        args.prompt, args.output,
        width_px=args.width, height_px=args.height,
        style=args.style, japanese_text=not args.no_japanese,
    )
    sys.exit(0 if result else 1)
