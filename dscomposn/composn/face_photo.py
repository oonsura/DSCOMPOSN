"""
ระบบตรวจสอบและปรับแต่งรูปภาพประจำตัวผู้ใช้ (Profile Photo Validator & Auto-Cropper)
- ตรวจจับใบหน้าด้วย OpenCV Haar Cascade
- ตรวจสอบว่าไม่ได้สวมแว่นดำ/แว่นกันแดด, ไม่ได้สวมหน้ากากอนามัย, และไม่ได้สวมหมวก
- ครอปเฉพาะบริเวณใบหน้า (อัตราส่วน 1:1 แบบรูปติดบัตร)
- ปรับขนาดและบีบอัดไฟล์ JPEG ให้มีขนาดไม่เกิน 1 MB (<= 1,000,000 bytes)
"""
import io
import os
import re

import cv2
import numpy as np
from django.core.files.base import ContentFile
from PIL import Image, ImageOps

MAX_RAW_UPLOAD_BYTES = 15 * 1024 * 1024   # 15 MB สูงสุดก่อนประมวลผล
MAX_OUTPUT_BYTES = 1_000_000              # ไม่เกิน 1 MB หลังปรับขนาด/บีบอัด
TARGET_SIZE = 600                         # ขนาดด้านกว้าง/สูงมาตรฐาน (600x600 px)

_CASCADES = {}


class FacePhotoError(Exception):
    """ข้อผิดพลาดจากการตรวจสอบรูปภาพใบหน้า"""


def _get_cascade(filename):
    if filename not in _CASCADES:
        path = os.path.join(cv2.data.haarcascades, filename)
        clf = cv2.CascadeClassifier(path)
        _CASCADES[filename] = clf if not clf.empty() else None
    return _CASCADES[filename]


def _ycrcb_skin_mask(ycrcb_img):
    """Mask พิกเซลสีผิวมนุษย์มาตรฐานในพื้นที่สี YCrCb (รองรับทุกโทนสีผิว)"""
    y = ycrcb_img[:, :, 0]
    cr = ycrcb_img[:, :, 1]
    cb = ycrcb_img[:, :, 2]
    return (y >= 38) & (cr >= 131) & (cr <= 178) & (cb >= 75) & (cb <= 130)


def _detect_primary_face(bgr):
    """ตรวจจับใบหน้าหลักในภาพ คืน (x, y, w, h) หรือ raise FacePhotoError"""
    h_img, w_img = bgr.shape[:2]
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    eq = cv2.equalizeHist(gray)
    min_dim = max(48, int(min(h_img, w_img) * 0.10))

    faces = ()
    for name, sf, mn in (
        ('haarcascade_frontalface_default.xml', 1.1, 5),
        ('haarcascade_frontalface_alt2.xml', 1.08, 4),
        ('haarcascade_frontalface_default.xml', 1.06, 3),
    ):
        clf = _get_cascade(name)
        if clf is None:
            continue
        det = clf.detectMultiScale(eq, scaleFactor=sf, minNeighbors=mn, minSize=(min_dim, min_dim))
        if len(det) > 0:
            faces = det
            break

    if len(faces) == 0:
        # ตรวจสอบกรณีสวมหน้ากากจนตรวจจับหน้าเต็มไม่พบ แต่ยังพบดวงตาทั้งสองข้าง + ใต้ตาไม่ใช่สีผิว
        eye_clf = _get_cascade('haarcascade_eye.xml')
        if eye_clf is not None:
            eyes = eye_clf.detectMultiScale(eq, scaleFactor=1.1, minNeighbors=5, minSize=(18, 18))
            if len(eyes) >= 2:
                eyes_sorted = sorted(eyes, key=lambda e: e[2] * e[3], reverse=True)[:4]
                for i in range(len(eyes_sorted)):
                    for j in range(i + 1, len(eyes_sorted)):
                        e1, e2 = eyes_sorted[i], eyes_sorted[j]
                        if abs(int(e1[1]) - int(e2[1])) < max(e1[3], e2[3]) * 0.7:
                            left_x = min(e1[0], e2[0])
                            right_x = max(e1[0] + e1[2], e2[0] + e2[2])
                            top_y = min(e1[1], e2[1])
                            span = right_x - left_x
                            if span >= 40:
                                my1 = min(h_img, int(top_y + span * 0.45))
                                my2 = min(h_img, int(top_y + span * 1.15))
                                mx1 = max(0, int(left_x))
                                mx2 = min(w_img, int(right_x))
                                if my2 - my1 > 15 and mx2 - mx1 > 15:
                                    roi = bgr[my1:my2, mx1:mx2]
                                    ycrcb = cv2.cvtColor(roi, cv2.COLOR_BGR2YCrCb)
                                    if float(np.mean(_ycrcb_skin_mask(ycrcb))) < 0.25:
                                        raise FacePhotoError(
                                            'ตรวจพบการสวมหน้ากากอนามัยหรือสิ่งปิดบังใบหน้าส่วนล่าง '
                                            'กรุณาถอดหน้ากากแล้วอัปโหลดรูปใหม่อีกครั้ง'
                                        )
        raise FacePhotoError(
            'ไม่พบใบหน้าชัดเจนในรูปภาพ กรุณาใช้รูปถ่ายหน้าตรง เห็นใบหน้าชัดเจน '
            'ไม่สวมหน้ากาก แว่นดำ หรือหมวก'
        )

    faces_sorted = sorted(faces, key=lambda f: int(f[2]) * int(f[3]), reverse=True)
    if len(faces_sorted) >= 2:
        a0 = int(faces_sorted[0][2]) * int(faces_sorted[0][3])
        a1 = int(faces_sorted[1][2]) * int(faces_sorted[1][3])
        if a1 > 0.55 * a0:
            raise FacePhotoError(
                'พบใบหน้ามากกว่า 1 คนในรูปภาพ กรุณาอัปโหลดรูปเดี่ยวที่เห็นใบหน้าของท่านเพียงคนเดียว'
            )

    fx, fy, fw, fh = [int(v) for v in faces_sorted[0]]
    return fx, fy, fw, fh


def _count_valid_eyes(gray_face, fw, fh):
    """นับจำนวนดวงตาที่ตรวจพบในครึ่งบนของใบหน้า (แยกฝั่งซ้าย/ขวา)"""
    upper = gray_face[int(fh * 0.14):int(fh * 0.56), :]
    if upper.size == 0:
        return 0
    eq_upper = cv2.equalizeHist(upper)
    found_left = False
    found_right = False
    for name in ('haarcascade_eye.xml', 'haarcascade_eye_tree_eyeglasses.xml'):
        clf = _get_cascade(name)
        if clf is None:
            continue
        eyes = clf.detectMultiScale(
            eq_upper,
            scaleFactor=1.08,
            minNeighbors=4,
            minSize=(max(12, int(fw * 0.10)), max(10, int(fh * 0.08))),
            maxSize=(int(fw * 0.44), int(fh * 0.36)),
        )
        for ex, ey, ew, eh in eyes:
            cx = ex + ew / 2.0
            if 0.12 * fw <= cx <= 0.48 * fw:
                found_left = True
            elif 0.52 * fw <= cx <= 0.88 * fw:
                found_right = True
        if found_left and found_right:
            return 2
    return int(found_left) + int(found_right)


def _check_sunglasses(bgr_face, hsv_face, ycrcb_face, gray_face):
    """ตรวจสอบว่าสวมแว่นดำ/แว่นกันแดดหรือไม่"""
    fh, fw = bgr_face.shape[:2]
    eye_y1, eye_y2 = int(fh * 0.24), int(fh * 0.46)
    left_x1, left_x2 = int(fw * 0.16), int(fw * 0.45)
    right_x1, right_x2 = int(fw * 0.55), int(fw * 0.84)

    cheek = hsv_face[int(fh * 0.50):int(fh * 0.65), int(fw * 0.22):int(fw * 0.78)]
    forehead = hsv_face[int(fh * 0.10):int(fh * 0.22), int(fw * 0.28):int(fw * 0.72)]
    left_eye_hsv = hsv_face[eye_y1:eye_y2, left_x1:left_x2]
    right_eye_hsv = hsv_face[eye_y1:eye_y2, right_x1:right_x2]
    left_eye_ycrcb = ycrcb_face[eye_y1:eye_y2, left_x1:left_x2]
    right_eye_ycrcb = ycrcb_face[eye_y1:eye_y2, right_x1:right_x2]

    if left_eye_hsv.size == 0 or right_eye_hsv.size == 0 or cheek.size == 0:
        return

    ref_v = float(np.median(cheek[:, :, 2]))
    if forehead.size > 0:
        ref_v = max(ref_v, float(np.median(forehead[:, :, 2])))

    v_l = left_eye_hsv[:, :, 2]
    v_r = right_eye_hsv[:, :, 2]
    s_l = left_eye_hsv[:, :, 1]
    s_r = right_eye_hsv[:, :, 1]

    med_vl = float(np.median(v_l))
    med_vr = float(np.median(v_r))
    dark_l = float(np.mean(v_l < 72))
    dark_r = float(np.mean(v_r < 72))
    sat_l = float(np.mean((s_l > 95) & (v_l > 40)))
    sat_r = float(np.mean((s_r > 95) & (v_r > 40)))
    skin_l = float(np.mean(_ycrcb_skin_mask(left_eye_ycrcb)))
    skin_r = float(np.mean(_ycrcb_skin_mask(right_eye_ycrcb)))

    valid_eyes = _count_valid_eyes(gray_face, fw, fh)

    # 1) เลนส์สีดำเข้มครอบคลุมบริเวณตาทั้งสองข้างชัดเจน (เข้มกว่าผิวแก้ม/หน้าผากมาก)
    if dark_l >= 0.45 and dark_r >= 0.45 and (ref_v - med_vl) >= 38 and (ref_v - med_vr) >= 38:
        if valid_eyes < 2 or (dark_l >= 0.60 and dark_r >= 0.60):
            raise FacePhotoError(
                'ตรวจพบการสวมแว่นดำ/แว่นกันแดด กรุณาถอดแว่นดำแล้วอัปโหลดรูปใหม่อีกครั้ง'
            )

    # 2) เลนส์ทึบมืดมากทั้งสองข้างและไม่พบดวงตา
    if med_vl < 58 and med_vr < 58 and ref_v >= 95 and valid_eyes == 0:
        raise FacePhotoError(
            'ตรวจพบการสวมแว่นดำ/แว่นกันแดด กรุณาถอดแว่นดำแล้วอัปโหลดรูปใหม่อีกครั้ง'
        )

    # 3) แว่นกันแดดเลนส์สีเข้ม/เลนส์ปรอท (สีสดผิดธรรมชาติและไม่เห็นสีผิวรอบดวงตา)
    if sat_l >= 0.42 and sat_r >= 0.42 and skin_l < 0.20 and skin_r < 0.20 and valid_eyes < 2:
        raise FacePhotoError(
            'ตรวจพบการสวมแว่นกันแดด/เลนส์สีทึบ กรุณาถอดแว่นกันแดดแล้วอัปโหลดรูปใหม่อีกครั้ง'
        )


def _check_mask(bgr_face, ycrcb_face, lab_face):
    """ตรวจสอบว่าสวมหน้ากากอนามัยหรือสิ่งปิดบังจมูก-ปากหรือไม่"""
    fh, fw = bgr_face.shape[:2]
    # บริเวณอ้างอิงผิวหน้าส่วนบน (หน้าผาก + สันจมูกระหว่างตา)
    forehead_y = ycrcb_face[int(fh * 0.12):int(fh * 0.24), int(fw * 0.28):int(fw * 0.72)]
    bridge_y = ycrcb_face[int(fh * 0.36):int(fh * 0.48), int(fw * 0.38):int(fw * 0.62)]
    forehead_lab = lab_face[int(fh * 0.12):int(fh * 0.24), int(fw * 0.28):int(fw * 0.72)]
    bridge_lab = lab_face[int(fh * 0.36):int(fh * 0.48), int(fw * 0.38):int(fw * 0.62)]

    # บริเวณจมูก-ปาก-คาง ที่หน้ากากจะปิดคลุม
    lower_y = ycrcb_face[int(fh * 0.56):int(fh * 0.90), int(fw * 0.22):int(fw * 0.78)]
    lower_lab = lab_face[int(fh * 0.56):int(fh * 0.90), int(fw * 0.22):int(fw * 0.78)]

    if lower_y.size == 0 or forehead_y.size == 0 or bridge_y.size == 0:
        return

    upper_skin_fh = float(np.mean(_ycrcb_skin_mask(forehead_y)))
    upper_skin_br = float(np.mean(_ycrcb_skin_mask(bridge_y)))
    upper_skin = max(upper_skin_fh, upper_skin_br)
    lower_skin = float(np.mean(_ycrcb_skin_mask(lower_y)))

    # เลือกบริเวณผิวหน้าส่วนบนที่เป็นผิวชัดเจนเป็นค่าอ้างอิง
    ref_y = bridge_y if upper_skin_br >= upper_skin_fh else forehead_y
    ref_lab = bridge_lab if upper_skin_br >= upper_skin_fh else forehead_lab

    ref_cr = float(np.median(ref_y[:, :, 1]))
    ref_cb = float(np.median(ref_y[:, :, 2]))
    low_cr = float(np.median(lower_y[:, :, 1]))
    low_cb = float(np.median(lower_y[:, :, 2]))
    ycrcb_dist = float(np.hypot(low_cr - ref_cr, low_cb - ref_cb))

    ref_l_med = np.median(ref_lab.reshape(-1, 3), axis=0)
    low_l_med = np.median(lower_lab.reshape(-1, 3), axis=0)
    lab_dist = float(np.linalg.norm(low_l_med - ref_l_med))

    # 1) ใบหน้าส่วนบนเป็นสีผิวชัดเจน แต่บริเวณจมูก-ปากแทบไม่มีสีผิวเลย (หน้ากากขาว/ฟ้า/ดำ/ผ้า)
    if upper_skin >= 0.40 and lower_skin < 0.28:
        raise FacePhotoError(
            'ตรวจพบการสวมหน้ากากอนามัยหรือสิ่งปิดบังใบหน้าส่วนล่าง '
            'กรุณาถอดหน้ากากแล้วอัปโหลดรูปใหม่อีกครั้ง'
        )

    # 2) สีผิวส่วนล่างต่ำมากพร้อมสีต่างจากใบหน้าส่วนบนอย่างชัดเจน
    if upper_skin >= 0.40 and lower_skin < 0.40 and (ycrcb_dist >= 16.0 or lab_dist >= 36.0):
        raise FacePhotoError(
            'ตรวจพบการสวมหน้ากากอนามัยหรือสิ่งปิดบังใบหน้าส่วนล่าง '
            'กรุณาถอดหน้ากากแล้วอัปโหลดรูปใหม่อีกครั้ง'
        )


def _check_hat(bgr, fx, fy, fw, fh, ycrcb_face, hsv_face, gray_face):
    """ตรวจสอบว่าสวมหมวกหรือสิ่งคลุมศีรษะหรือไม่ (แยกแยะพื้นหลังและผมธรรมชาติออก)"""
    h_img, w_img = bgr.shape[:2]

    # 1) วิเคราะห์โซนเหนือหน้าผาก (Crown band) โดยตัดพิกเซลพื้นหลังออก
    cy1 = max(0, fy - int(fh * 0.40))
    cy2 = min(h_img, fy + int(fh * 0.04))
    cx1 = max(0, fx + int(fw * 0.08))
    cx2 = min(w_img, fx + int(fw * 0.92))

    if cy2 - cy1 < 12 or cx2 - cx1 < 20:
        return

    crown_bgr = bgr[cy1:cy2, cx1:cx2]
    crown_hsv = cv2.cvtColor(crown_bgr, cv2.COLOR_BGR2HSV)
    crown_ycrcb = cv2.cvtColor(crown_bgr, cv2.COLOR_BGR2YCrCb)
    crown_lab = cv2.cvtColor(crown_bgr, cv2.COLOR_BGR2LAB).astype(np.float32)

    # เก็บตัวอย่างสีพื้นหลังจากมุมซ้ายบนและขวาบนนอกกรอบศีรษะ
    bg_samples = []
    bg_y2 = max(10, min(h_img, fy + int(fh * 0.15)))
    if fx > int(fw * 0.15):
        left_bg = bgr[0:bg_y2, 0:max(5, fx - int(fw * 0.10))]
        if left_bg.size > 0:
            bg_samples.append(cv2.cvtColor(left_bg, cv2.COLOR_BGR2LAB).reshape(-1, 3))
    if fx + fw + int(fw * 0.15) < w_img:
        right_bg = bgr[0:bg_y2, min(w_img - 5, fx + fw + int(fw * 0.10)):w_img]
        if right_bg.size > 0:
            bg_samples.append(cv2.cvtColor(right_bg, cv2.COLOR_BGR2LAB).reshape(-1, 3))

    non_bg_mask = np.ones(crown_bgr.shape[:2], dtype=bool)
    if bg_samples:
        bg_all = np.vstack(bg_samples).astype(np.float32)
        bg_med = np.median(bg_all, axis=0)
        dist_bg = np.linalg.norm(crown_lab - bg_med, axis=2)
        non_bg_mask = dist_bg > 22.0

    non_skin_mask = ~_ycrcb_skin_mask(crown_ycrcb)
    fg_mask = non_bg_mask & non_skin_mask

    # A) หมวกสีสด/มีสีสันชัดเจน (สีแดง น้ำเงิน เขียว เหลือง ม่วง ฯลฯ ที่ไม่ใช่สีพื้นหลังและไม่ใช่ผมธรรมชาติ)
    # ผมธรรมชาติมี S ต่ำ (ดำ/น้ำตาลเข้ม) หรือหากเป็นน้ำตาลจะมี Hue โทนอุ่นไม่สด (S < 80)
    sat = crown_hsv[:, :, 1]
    val = crown_hsv[:, :, 2]
    vivid_hat_mask = fg_mask & (sat >= 85) & (val >= 55)
    vivid_ratio = float(np.mean(vivid_hat_mask))
    if vivid_ratio >= 0.28:
        raise FacePhotoError(
            'ตรวจพบการสวมหมวกหรือสิ่งคลุมศีรษะ กรุณาถอดหมวกแล้วอัปโหลดรูปใหม่อีกครั้ง'
        )

    # B) หมวกสีขาว/สีอ่อนทึบเรียบเหนือหน้าผาก (ไม่ใช่สีพื้นหลัง)
    white_hat_mask = fg_mask & (val >= 195) & (sat <= 38)
    if float(np.mean(white_hat_mask)) >= 0.34:
        gray_crown = cv2.cvtColor(crown_bgr, cv2.COLOR_BGR2GRAY)
        lap_var = float(cv2.Laplacian(gray_crown, cv2.CV_64F).var())
        if lap_var < 260.0:
            raise FacePhotoError(
                'ตรวจพบการสวมหมวกหรือสิ่งคลุมศีรษะ กรุณาถอดหมวกแล้วอัปโหลดรูปใหม่อีกครั้ง'
            )

    # C) ปีกหมวกแก๊ป/หมวกสีเข้มหรือสีขาวพาดขวางหน้าผาก (ขอบแนวนอนคมชัด + หน้าผากถูกปิดแทบไม่เห็นผิว)
    fh_strip_y = ycrcb_face[0:int(fh * 0.20), int(fw * 0.20):int(fw * 0.80)]
    fh_strip_gray = gray_face[0:int(fh * 0.22), int(fw * 0.18):int(fw * 0.82)]
    if fh_strip_y.size > 0 and fh_strip_gray.size > 0:
        forehead_skin = float(np.mean(_ycrcb_skin_mask(fh_strip_y)))
        if forehead_skin < 0.24:
            gx = cv2.Sobel(fh_strip_gray, cv2.CV_32F, 1, 0, ksize=3)
            gy = cv2.Sobel(fh_strip_gray, cv2.CV_32F, 0, 1, ksize=3)
            mean_gx = float(np.mean(np.abs(gx))) + 1e-5
            mean_gy = float(np.mean(np.abs(gy)))
            # ปีกหมวกมีขอบแนวนอนเด่นมาก (gy >> gx) ต่างจากผมหน้าม้าที่มีเส้นแนวตั้ง (gx สูง)
            if mean_gy >= 3.0 and (mean_gy / mean_gx) >= 2.2:
                raise FacePhotoError(
                    'ตรวจพบการสวมหมวกหรือปีกหมวกปิดบังหน้าผาก กรุณาถอดหมวกแล้วอัปโหลดรูปใหม่อีกครั้ง'
                )


def process_and_validate_profile_photo(uploaded_file, user_code='user'):
    """
    ตรวจสอบรูปภาพใบหน้า (ห้ามสวมหน้ากาก / แว่นดำ / หมวก), ครอปเฉพาะใบหน้า 1:1,
    ย่อขนาดและบีบอัด JPEG ให้ไม่เกิน 1 MB
    คืนค่า (ContentFile, info_dict) หรือ raise FacePhotoError
    """
    if not uploaded_file:
        raise FacePhotoError('กรุณาเลือกไฟล์รูปภาพ')
    raw_size = getattr(uploaded_file, 'size', None)
    if raw_size is not None and raw_size > MAX_RAW_UPLOAD_BYTES:
        raise FacePhotoError('ไฟล์รูปภาพต้นฉบับมีขนาดใหญ่เกิน 15 MB กรุณาลดขนาดก่อนอัปโหลด')

    try:
        uploaded_file.seek(0)
        raw_bytes = uploaded_file.read()
        if not raw_bytes:
            raise FacePhotoError('ไฟล์รูปภาพว่างเปล่า')
        if len(raw_bytes) > MAX_RAW_UPLOAD_BYTES:
            raise FacePhotoError('ไฟล์รูปภาพต้นฉบับมีขนาดใหญ่เกิน 15 MB')
        pil_img = Image.open(io.BytesIO(raw_bytes))
        pil_img = ImageOps.exif_transpose(pil_img)
        if pil_img.mode != 'RGB':
            pil_img = pil_img.convert('RGB')
    except FacePhotoError:
        raise
    except Exception:
        raise FacePhotoError('ไฟล์ที่อัปโหลดไม่ใช่รูปภาพที่ถูกต้อง (รองรับ JPG, PNG, WEBP)')

    orig_w, orig_h = pil_img.size
    if orig_w < 100 or orig_h < 100:
        raise FacePhotoError('ความละเอียดรูปภาพต่ำเกินไป (ต้องกว้างและสูงอย่างน้อย 100 พิกเซล)')

    # ย่อขนาดชั่วคราวสำหรับการตรวจจับใบหน้า (ไม่เกิน 1200 px) เพื่อความแม่นยำและรวดเร็ว
    scale = 1.0
    max_side = max(orig_w, orig_h)
    if max_side > 1200:
        scale = 1200.0 / max_side
        det_pil = pil_img.resize((int(orig_w * scale), int(orig_h * scale)), Image.Resampling.LANCZOS)
    else:
        det_pil = pil_img

    rgb = np.array(det_pil)
    bgr = cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)

    # 1. ตรวจจับใบหน้า
    fx, fy, fw, fh = _detect_primary_face(bgr)

    bgr_face = bgr[fy:fy + fh, fx:fx + fw]
    hsv_face = cv2.cvtColor(bgr_face, cv2.COLOR_BGR2HSV)
    ycrcb_face = cv2.cvtColor(bgr_face, cv2.COLOR_BGR2YCrCb)
    lab_face = cv2.cvtColor(bgr_face, cv2.COLOR_BGR2LAB).astype(np.float32)
    gray_face = cv2.cvtColor(bgr_face, cv2.COLOR_BGR2GRAY)

    # 2. ตรวจสอบแว่นดำ / หน้ากากอนามัย / หมวก
    _check_sunglasses(bgr_face, hsv_face, ycrcb_face, gray_face)
    _check_mask(bgr_face, ycrcb_face, lab_face)
    _check_hat(bgr, fx, fy, fw, fh, ycrcb_face, hsv_face, gray_face)

    # 3. คำนวณกรอบครอปเฉพาะใบหน้า 1:1 บนภาพต้นฉบับความละเอียดเต็ม
    inv = 1.0 / scale
    ofx, ofy, ofw, ofh = fx * inv, fy * inv, fw * inv, fh * inv
    cx = ofx + ofw / 2.0
    cy = ofy + ofh * 0.44  # ขยับขึ้นเล็กน้อยเพื่อเก็บศีรษะและคางได้สมดุล
    side = max(ofw, ofh) * 1.80
    side = min(side, float(orig_w), float(orig_h))

    half = side / 2.0
    x1 = int(round(max(0.0, min(orig_w - side, cx - half))))
    y1 = int(round(max(0.0, min(orig_h - side, cy - half))))
    x2 = min(orig_w, x1 + int(round(side)))
    y2 = min(orig_h, y1 + int(round(side)))

    cropped = pil_img.crop((x1, y1, x2, y2))
    if cropped.width > TARGET_SIZE or cropped.height > TARGET_SIZE:
        cropped = cropped.resize((TARGET_SIZE, TARGET_SIZE), Image.Resampling.LANCZOS)

    # 4. บีบอัดเป็น JPEG ขนาดไม่เกิน 1 MB (<= 1,000,000 bytes)
    out_bytes = None
    cur_img = cropped
    for target_dim in (TARGET_SIZE, 480, 360):
        if cur_img.width > target_dim:
            cur_img = cur_img.resize((target_dim, target_dim), Image.Resampling.LANCZOS)
        for q in (88, 82, 75, 68, 60, 50):
            buf = io.BytesIO()
            cur_img.save(buf, format='JPEG', quality=q, optimize=True, progressive=True)
            data = buf.getvalue()
            if len(data) <= MAX_OUTPUT_BYTES:
                out_bytes = data
                break
        if out_bytes is not None:
            break

    if out_bytes is None:
        raise FacePhotoError('ไม่สามารถบีบอัดรูปภาพให้มีขนาดไม่เกิน 1 MB ได้ กรุณาใช้รูปอื่น')

    safe_code = re.sub(r'[^a-zA-Z0-9_-]', '_', str(user_code or 'user'))[:30]
    content_file = ContentFile(out_bytes, name=f'face_{safe_code}.jpg')
    return content_file, {
        'size_bytes': len(out_bytes),
        'width': cur_img.width,
        'height': cur_img.height,
        'face_box': (int(ofx), int(ofy), int(ofw), int(ofh)),
    }
