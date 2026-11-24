# dscomposn

ระบบรวม **COMPOSN** (สิทธิ์ค่าย) + **Food** (เวลาเรียน / คะแนน / สั่งอาหาร) — Django 4.2

## เริ่มใช้งาน

```powershell
pip install -r requirements.txt
python manage.py migrate        # สร้าง DB + Fixtures (admin/1234, เหตุผลขาดเรียน, สีเตือนภัย, ค่าเริ่มต้น)
python manage.py runserver
```

เปิด http://127.0.0.1:8000/ — เข้าสู่ระบบด้วย `admin` / `1234` (ต้องเปลี่ยนรหัสผ่านครั้งแรก)

- Member: Username = รหัสค่าย 1 (เช่น `69DS01`) หรือรหัสค่าย 2 (ถ้ามี เช่น `68KM01`), รหัสผ่านเริ่มต้น = รหัสค่าย 1 เต็ม
- Django Admin อยู่ที่ `/django-admin/`
- ทดสอบ: `python manage.py test` (41 tests)

## ตัวแปรสภาพแวดล้อม (.env)

| ตัวแปร | ค่าเริ่มต้น | หมายเหตุ |
|---|---|---|
| `DJANGO_SECRET_KEY` | dev key | **ต้องเปลี่ยนใน production** |
| `DJANGO_DEBUG` | True | production = False |
| `DJANGO_ALLOWED_HOSTS` | * | |
| `DATABASE_ENGINE` | sqlite | `postgresql` + `DATABASE_NAME/USER/PASSWORD/HOST/PORT` |
| `AUTO_LOGOUT_SECONDS` | 300 | Auto Logout (5 นาที) |
| `DJANGO_SECURE_COOKIES` | False | production (HTTPS) = True |

## โครงสร้าง

- `composn/` — User, สิทธิ์ (`rights.py`), ระดับชั้น (`levels.py`), Excel/Backup (`excel.py`), Audit (`audit.py`), middleware
- `food/` — เวลาเรียน, คะแนน, สั่งอาหาร
- `backups/` — ไฟล์สำรอง Excel, `audit_logs/` — ไฟล์ Audit รายวัน

## Deploy (สรุป)

`DJANGO_DEBUG=False`, ตั้ง `SECRET_KEY`/`ALLOWED_HOSTS`/`DATABASE_*`, `python manage.py collectstatic`,
รันด้วย gunicorn/waitress หลัง reverse proxy (HTTPS) และเสิร์ฟ `/static/` + `/media/` ผ่าน web server
(ต้องติดตั้ง WhiteNoise หรือใช้ nginx/IIS สำหรับ static)
