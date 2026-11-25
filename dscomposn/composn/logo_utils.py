"""
เครื่องมือแปลงรูปตราสัญลักษณ์ (Logos / Emblems) ให้มีพื้นหลังโปร่งใส (RGBA PNG)
สำหรับใช้ทั้งการแสดงผลบนหน้าเว็บและเอกสาร PDF
"""
from collections import deque
import io
import os

from django.core.files.base import ContentFile
from PIL import Image, ImageOps

LOGO_FIELD_NAMES = (
    'login_logo',
    'logo_school',
    'logo_posn',
    'logo_center',
    'logo_obec',
    'logo_garuda',
    'site_favicon',
)


def _has_transparent_border(rgba_img):
    """ตรวจสอบว่าพิกเซลขอบรูปมีความโปร่งใสอยู่แล้วหรือไม่ (alpha < 250)"""
    w, h = rgba_img.size
    if w <= 0 or h <= 0:
        return False
    px = rgba_img.load()
    for x in range(w):
        if px[x, 0][3] < 250 or px[x, h - 1][3] < 250:
            return True
    for y in range(h):
        if px[0, y][3] < 250 or px[w - 1, y][3] < 250:
            return True
    return False


def _chebyshev_dist(rgb1, rgb2):
    return max(abs(rgb1[0] - rgb2[0]), abs(rgb1[1] - rgb2[1]), abs(rgb1[2] - rgb2[2]))


def _channel_median(rgb_list):
    n = len(rgb_list)
    rs = sorted(c[0] for c in rgb_list)
    gs = sorted(c[1] for c in rgb_list)
    bs = sorted(c[2] for c in rgb_list)
    mid = n // 2
    if n % 2 == 1:
        return (rs[mid], gs[mid], bs[mid])
    return (
        (rs[mid - 1] + rs[mid]) // 2,
        (gs[mid - 1] + gs[mid]) // 2,
        (bs[mid - 1] + bs[mid]) // 2,
    )


def _is_monochrome_black_on_white(rgba_img):
    """ตรวจสอบว่าเป็นตราขาว-ดำ / ภาพลายเส้นบนพื้นขาวหรือไม่ (เช่น ตราครุฑ)"""
    w, h = rgba_img.size
    px = rgba_img.load()
    corners = [px[0, 0][:3], px[w - 1, 0][:3], px[0, h - 1][:3], px[w - 1, h - 1][:3]]
    if sum(1 for r, g, b in corners if r >= 230 and g >= 230 and b >= 230) < 3:
        return False

    data = list(rgba_img.getdata())
    total = len(data)
    step = max(1, total // 4000)
    sampled = data[::step]
    gray_cnt = 0
    min_lum = 255.0
    for r, g, b, _a in sampled:
        if max(r, g, b) - min(r, g, b) <= 15:
            gray_cnt += 1
        lum = 0.299 * r + 0.587 * g + 0.114 * b
        if lum < min_lum:
            min_lum = lum
    return (gray_cnt / max(1, len(sampled))) >= 0.98 and min_lum < 120.0


def _monochrome_luminance_to_alpha(rgba_img):
    """แปลงภาพขาว-ดำเป็นสีดำโปร่งใสตามค่าความสว่าง (luminance-to-alpha)"""
    new_pixels = []
    for r, g, b, a in rgba_img.getdata():
        if a == 0:
            new_pixels.append((0, 0, 0, 0))
            continue
        lum = 0.299 * r + 0.587 * g + 0.114 * b
        if lum >= 242.0:
            new_pixels.append((0, 0, 0, 0))
        elif lum <= 40.0:
            new_pixels.append((0, 0, 0, a))
        else:
            alpha = int(round((242.0 - lum) * 255.0 / 202.0))
            alpha = max(0, min(255, alpha))
            if a < 255:
                alpha = min(a, alpha)
            new_pixels.append((0, 0, 0, alpha))
    out = Image.new('RGBA', rgba_img.size)
    out.putdata(new_pixels)
    return out


def _uniform_corner_bg(rgba_img):
    """หาสีพื้นหลังตัวแทนจากมุมทั้ง 4 ของภาพ (หากมุมไม่สม่ำเสมอคืนค่า None)"""
    w, h = rgba_img.size
    px = rgba_img.load()
    corners = [px[0, 0][:3], px[w - 1, 0][:3], px[0, h - 1][:3], px[w - 1, h - 1][:3]]
    near_white = [c for c in corners if c[0] >= 230 and c[1] >= 230 and c[2] >= 230]
    if len(near_white) >= 3:
        return _channel_median(near_white)
    med = _channel_median(corners)
    if sum(1 for c in corners if _chebyshev_dist(c, med) <= 20) >= 3:
        return med
    return None


def _flood_remove_exterior_bg(rgba_img, bg_rgb, tol=32, feather_tol=68):
    """ลบพื้นหลังภายนอกด้วย 4-connected BFS จากขอบภาพ พร้อมปรับขอบนุ่ม (feather)"""
    w, h = rgba_img.size
    px = rgba_img.load()
    visited = bytearray(w * h)
    queue = deque()

    def try_seed(x, y):
        idx = y * w + x
        if visited[idx]:
            return
        r, g, b, _a = px[x, y]
        if _chebyshev_dist((r, g, b), bg_rgb) <= tol:
            visited[idx] = 1
            queue.append((x, y))

    for x in range(w):
        try_seed(x, 0)
        try_seed(x, h - 1)
    for y in range(h):
        try_seed(0, y)
        try_seed(w - 1, y)

    while queue:
        cx, cy = queue.popleft()
        r, g, b, _a = px[cx, cy]
        px[cx, cy] = (r, g, b, 0)
        if cx > 0:
            nx, ny = cx - 1, cy
            nidx = ny * w + nx
            if not visited[nidx] and _chebyshev_dist(px[nx, ny][:3], bg_rgb) <= tol:
                visited[nidx] = 1
                queue.append((nx, ny))
        if cx + 1 < w:
            nx, ny = cx + 1, cy
            nidx = ny * w + nx
            if not visited[nidx] and _chebyshev_dist(px[nx, ny][:3], bg_rgb) <= tol:
                visited[nidx] = 1
                queue.append((nx, ny))
        if cy > 0:
            nx, ny = cx, cy - 1
            nidx = ny * w + nx
            if not visited[nidx] and _chebyshev_dist(px[nx, ny][:3], bg_rgb) <= tol:
                visited[nidx] = 1
                queue.append((nx, ny))
        if cy + 1 < h:
            nx, ny = cx, cy + 1
            nidx = ny * w + nx
            if not visited[nidx] and _chebyshev_dist(px[nx, ny][:3], bg_rgb) <= tol:
                visited[nidx] = 1
                queue.append((nx, ny))

    # Feather ขอบพิกเซลที่ติดกับพื้นหลังที่ถูกลบ เพื่อลดขอบหยัก/ขอบขาว
    span = float(feather_tol - tol)
    feathered = set()
    for idx, v in enumerate(visited):
        if not v:
            continue
        cx = idx % w
        cy = idx // w
        for nx, ny in ((cx - 1, cy), (cx + 1, cy), (cx, cy - 1), (cx, cy + 1)):
            if 0 <= nx < w and 0 <= ny < h:
                nidx = ny * w + nx
                if not visited[nidx] and nidx not in feathered:
                    feathered.add(nidx)
                    r, g, b, old_a = px[nx, ny]
                    d = _chebyshev_dist((r, g, b), bg_rgb)
                    if tol < d <= feather_tol:
                        scaled_a = int(round((d - tol) * 255.0 / span))
                        px[nx, ny] = (r, g, b, min(old_a, max(0, min(255, scaled_a))))
    return rgba_img


def remove_logo_background(pil_image):
    """
    แปลงรูปตราสัญลักษณ์ให้มีพื้นหลังโปร่งใส (RGBA):
    - หากขอบภาพโปร่งใสอยู่แล้ว คืนค่า RGBA เดิมโดยไม่แตะรายละเอียดสีขาวภายใน
    - หากเป็นภาพขาว-ดำ (ลายเส้นบนพื้นขาว) ใช้ luminance-to-alpha
    - หากเป็นภาพสีพื้นหลังทึบ ใช้ BFS ลบเฉพาะพื้นหลังภายนอกจากขอบภาพ + feather ขอบ
    """
    try:
        img = ImageOps.exif_transpose(pil_image)
    except Exception:
        img = pil_image

    rgba = img.convert('RGBA')
    w, h = rgba.size
    if w <= 2 and h <= 2:
        return rgba

    max_logo_dim = 400
    if max(w, h) > max_logo_dim:
        resampling = getattr(Image, 'Resampling', Image).LANCZOS
        scale = float(max_logo_dim) / float(max(w, h))
        new_size = (max(1, int(round(w * scale))), max(1, int(round(h * scale))))
        rgba = rgba.resize(new_size, resampling)

    if _has_transparent_border(rgba):
        return rgba

    if _is_monochrome_black_on_white(rgba):
        return _monochrome_luminance_to_alpha(rgba)

    bg_rgb = _uniform_corner_bg(rgba)
    if bg_rgb is None:
        return rgba

    return _flood_remove_exterior_bg(rgba, bg_rgb)


def _stem_from_filename(name, default_stem='logo'):
    base = os.path.basename(name or '')
    stem = os.path.splitext(base)[0].strip()
    return stem or default_stem


def _encode_compact_rgba_png(rgba_img, max_dim=400):
    """ย่อขนาดและบีบอัด PNG พื้นหลังโปร่งใส (โหมด RGBA) ให้มีขนาดกะทัดรัดโดยคงค่าสีและขอบโปร่งใสครบถ้วน"""
    if rgba_img.mode != 'RGBA':
        rgba_img = rgba_img.convert('RGBA')
    w, h = rgba_img.size
    if max(w, h) > max_dim:
        resampling = getattr(Image, 'Resampling', Image).LANCZOS
        scale = float(max_dim) / float(max(w, h))
        new_size = (max(1, int(round(w * scale))), max(1, int(round(h * scale))))
        rgba_img = rgba_img.resize(new_size, resampling)

    buf_rgba = io.BytesIO()
    rgba_img.save(buf_rgba, format='PNG', optimize=True, compress_level=9)
    return buf_rgba.getvalue()


def process_logo_upload(uploaded_file, is_favicon=False):
    """แปลงไฟล์รูปตราที่อัปโหลดให้เป็นไฟล์ PNG พื้นหลังโปร่งใสขนาดกะทัดรัด (ContentFile)"""
    if not uploaded_file:
        return uploaded_file
    try:
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        with Image.open(uploaded_file) as img:
            img.load()
            out_rgba = remove_logo_background(img)
        max_dim = 64 if is_favicon else 400
        png_bytes = _encode_compact_rgba_png(out_rgba, max_dim=max_dim)
        stem = _stem_from_filename(getattr(uploaded_file, 'name', 'logo.png'))
        cf = ContentFile(png_bytes, name=f'{stem}.png')
        cf._already_compressed = True
        return cf
    except Exception:
        try:
            if hasattr(uploaded_file, 'seek'):
                uploaded_file.seek(0)
        except Exception:
            pass
        return uploaded_file


def convert_existing_setting_logos_to_transparent(setting):
    """แปลงและย่อขนาดรูปตราสัญลักษณ์ทั้งหมดที่บันทึกอยู่ใน SystemSetting ให้เป็น PNG พื้นหลังโปร่งใสขนาดกะทัดรัด"""
    converted = []
    for field_name in LOGO_FIELD_NAMES:
        fld = getattr(setting, field_name, None)
        if not fld:
            continue
        try:
            old_path = fld.path
        except Exception:
            continue
        if not old_path or not os.path.exists(old_path):
            continue
        max_dim = 64 if field_name == 'site_favicon' else 400
        try:
            with Image.open(old_path) as im:
                im.load()
                orig_mode = im.mode
                orig_w, orig_h = im.size
                rgba_check = im.convert('RGBA')
                needs_conv = (
                    (not fld.name.lower().endswith('.png'))
                    or (orig_mode != 'RGBA')
                    or (not _has_transparent_border(rgba_check))
                    or (max(orig_w, orig_h) > max_dim)
                    or (os.path.getsize(old_path) > (20_000 if field_name == 'site_favicon' else 80_000))
                )
                if not needs_conv:
                    continue
                out_rgba = remove_logo_background(im)
        except Exception:
            continue

        png_bytes = _encode_compact_rgba_png(out_rgba, max_dim=max_dim)
        stem = _stem_from_filename(fld.name, default_stem=field_name)
        fld.save(f'{stem}.png', ContentFile(png_bytes), save=False)
        try:
            new_path = fld.path
            if old_path != new_path and os.path.exists(old_path):
                os.remove(old_path)
        except Exception:
            pass
        converted.append(field_name)

    if converted:
        setting.save()
    return converted
