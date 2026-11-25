"""
เครื่องมือย่อขนาดและบีบอัดรูปภาพอัตโนมัติ (Image Compression & Orphan Cleanup Utilities)
เพื่อประหยัดพื้นที่จัดเก็บและเพิ่มความเร็วในการโหลดหน้าเว็บเมื่อใช้งานพร้อมกันหลายคน
"""
import io
import os
import re

from django.core.files.base import ContentFile
from PIL import Image, ImageOps

MAX_UPLOAD_RAW_BYTES = 20 * 1024 * 1024  # จำกัดไฟล์ต้นฉบับสูงสุด 20 MB

FOOD_MENU_MAX_DIM = 640
FOOD_MENU_MAX_BYTES = 90_000       # <= 90 KB ต่อรูปเมนูอาหาร (แสดงผลจริง ~280x145 px)

FOOD_ORDER_MAX_DIM = 960
FOOD_ORDER_MAX_BYTES = 150_000     # <= 150 KB ต่อรูปภาพรวมร้านอาหาร

CERT_BG_MAX_DIM = 2000
CERT_BG_MAX_BYTES = 350_000        # <= 350 KB ต่อรูปพื้นหลังเกียรติบัตร A4 แนวนอน


class ImageUploadError(ValueError):
    """ข้อผิดพลาดเมื่อไฟล์รูปภาพไม่ถูกต้องหรือมีขนาดใหญ่เกินกำหนด"""


def _safe_stem(filename, default_stem='image'):
    base = os.path.basename(str(filename or ''))
    stem = os.path.splitext(base)[0].strip()
    # อนุญาตทั้งอักษรไทย อังกฤษ ตัวเลข ขีดล่าง และขีดกลาง
    cleaned = re.sub(r'[^\w\-ก-๙]', '_', stem)[:60].strip('_')
    return cleaned or default_stem


def compress_photo_upload(
    uploaded_file,
    max_dim=800,
    max_bytes=120_000,
    default_stem='image',
    qualities=(82, 76, 70, 62, 55, 48),
):
    """ย่อขนาดและบีบอัดรูปภาพให้เป็น Progressive JPEG ขนาดเล็กแต่คงความคมชัดตามการใช้งานจริง"""
    if not uploaded_file:
        return None
    try:
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)
        raw = uploaded_file.read()
    except Exception as exc:
        raise ImageUploadError('ไม่สามารถอ่านไฟล์รูปภาพได้') from exc

    if not raw:
        raise ImageUploadError('ไฟล์รูปภาพว่างเปล่า')
    if len(raw) > MAX_UPLOAD_RAW_BYTES:
        raise ImageUploadError('ไฟล์รูปภาพต้นฉบับมีขนาดใหญ่เกิน 20 MB')

    try:
        with Image.open(io.BytesIO(raw)) as src:
            img = ImageOps.exif_transpose(src)
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                rgba = img.convert('RGBA')
                bg = Image.new('RGBA', rgba.size, (255, 255, 255, 255))
                bg.paste(rgba, mask=rgba.split()[3])
                rgb = bg.convert('RGB')
            elif img.mode != 'RGB':
                rgb = img.convert('RGB')
            else:
                rgb = img.copy()
    except Exception as exc:
        raise ImageUploadError('ไฟล์ที่อัปโหลดไม่ใช่รูปภาพที่ถูกต้อง (รองรับ JPG, PNG, WEBP)') from exc

    resampling = getattr(Image, 'Resampling', Image).LANCZOS
    dim_steps = [
        max_dim,
        max(240, int(round(max_dim * 0.82))),
        max(200, int(round(max_dim * 0.65))),
    ]

    best_bytes = None
    cur = rgb
    for target_dim in dim_steps:
        w, h = cur.size
        if max(w, h) > target_dim:
            scale = float(target_dim) / float(max(w, h))
            new_size = (max(1, int(round(w * scale))), max(1, int(round(h * scale))))
            cur = cur.resize(new_size, resampling)
        for q in qualities:
            buf = io.BytesIO()
            cur.save(buf, format='JPEG', quality=q, optimize=True, progressive=True)
            data = buf.getvalue()
            best_bytes = data
            if len(data) <= max_bytes:
                stem = _safe_stem(getattr(uploaded_file, 'name', ''), default_stem=default_stem)
                cf = ContentFile(data, name=f'{stem}.jpg')
                cf._already_compressed = True
                return cf

    stem = _safe_stem(getattr(uploaded_file, 'name', ''), default_stem=default_stem)
    cf = ContentFile(best_bytes, name=f'{stem}.jpg')
    cf._already_compressed = True
    return cf


def compress_profile_fallback_image(uploaded_file):
    """ย่อและบีบอัดรูปโปรไฟล์กรณีอัปโหลดตรงผ่าน Model/Django Admin (สูงสุด 420px, ไม่เกิน ~100 KB)"""
    return compress_photo_upload(
        uploaded_file,
        max_dim=420,
        max_bytes=100_000,
        default_stem='profile',
        qualities=(82, 75, 68, 60, 52, 44),
    )


def ensure_compressed_field(instance, field_name, compressor_fn):
    """รับประกันว่า ImageField ที่เพิ่งอัปโหลดใหม่ (_committed=False) จะถูกย่อและบีบอัดอัตโนมัติใน Model.save() เสมอ
    โดยไม่บีบอัดซ้ำหากผ่านตัวบีบอัดจาก Form/View มาแล้ว (_already_compressed=True)"""
    fld = getattr(instance, field_name, None)
    if not fld or getattr(fld, '_committed', True):
        return
    inner_file = getattr(fld, 'file', None)
    if getattr(fld, '_already_compressed', False) or getattr(inner_file, '_already_compressed', False):
        return
    try:
        compressed = compressor_fn(fld)
    except Exception:
        return
    if compressed is not None:
        compressed._already_compressed = True
        setattr(instance, field_name, compressed)


def compress_food_menu_image(uploaded_file):
    """ย่อและบีบอัดรูปเมนูอาหาร (สูงสุด 640px, ไม่เกิน ~90 KB)"""
    return compress_photo_upload(
        uploaded_file,
        max_dim=FOOD_MENU_MAX_DIM,
        max_bytes=FOOD_MENU_MAX_BYTES,
        default_stem='menu',
        qualities=(80, 74, 68, 60, 52, 45),
    )


def compress_food_order_image(uploaded_file):
    """ย่อและบีบอัดรูปภาพรวมรายการอาหาร (สูงสุด 960px, ไม่เกิน ~150 KB)"""
    return compress_photo_upload(
        uploaded_file,
        max_dim=FOOD_ORDER_MAX_DIM,
        max_bytes=FOOD_ORDER_MAX_BYTES,
        default_stem='order',
        qualities=(82, 75, 68, 60, 52, 45),
    )


def compress_certificate_bg_image(uploaded_file):
    """ย่อและบีบอัดรูปพื้นหลังเกียรติบัตร A4 แนวนอน (สูงสุด 2000px, ไม่เกิน ~350 KB)"""
    return compress_photo_upload(
        uploaded_file,
        max_dim=CERT_BG_MAX_DIM,
        max_bytes=CERT_BG_MAX_BYTES,
        default_stem='cert_bg',
        qualities=(86, 80, 74, 68, 60, 52),
    )


def delete_file_if_unreferenced(field_file, model_cls, field_name, exclude_pk=None, storage=None):
    """ลบไฟล์รูปเก่าออกจากดิสก์เฉพาะกรณีไม่มีเรคคอร์ดอื่นอ้างอิงไฟล์ชื่อเดียวกันอยู่ (รองรับการคัดลอกรายการอาหาร)"""
    if not field_file:
        return False
    if isinstance(field_file, str):
        file_name = field_file
        if storage is None:
            from django.core.files.storage import default_storage
            storage = default_storage
    else:
        file_name = getattr(field_file, 'name', '') or ''
        storage = storage or getattr(field_file, 'storage', None)
    if not file_name or storage is None:
        return False
    qs = model_cls._default_manager.filter(**{field_name: file_name})
    if exclude_pk is not None:
        qs = qs.exclude(pk=exclude_pk)
    if qs.exists():
        return False
    try:
        if storage.exists(file_name):
            storage.delete(file_name)
            return True
    except Exception:
        pass
    return False
