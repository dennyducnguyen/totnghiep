# -*- coding: utf-8 -*-
"""Danh muc ten chuan: 63 tinh thanh (truoc sap nhap 2025), 6 vung, ca nuoc.
Kem ham chuan hoa ten de so khop bat chap bien the chinh ta (Hoà/Hòa, TP./Tp, ...).
"""
import re
import unicodedata

PROVINCES_63 = [
    # Dong bang song Hong (11)
    'Hà Nội', 'Vĩnh Phúc', 'Bắc Ninh', 'Quảng Ninh', 'Hải Dương', 'Hải Phòng',
    'Hưng Yên', 'Thái Bình', 'Hà Nam', 'Nam Định', 'Ninh Bình',
    # Trung du va mien nui phia Bac (14)
    'Hà Giang', 'Cao Bằng', 'Bắc Kạn', 'Tuyên Quang', 'Lào Cai', 'Yên Bái',
    'Thái Nguyên', 'Lạng Sơn', 'Bắc Giang', 'Phú Thọ', 'Điện Biên', 'Lai Châu',
    'Sơn La', 'Hòa Bình',
    # Bac Trung Bo va Duyen hai mien Trung (14)
    'Thanh Hóa', 'Nghệ An', 'Hà Tĩnh', 'Quảng Bình', 'Quảng Trị',
    'Thừa Thiên - Huế', 'Đà Nẵng', 'Quảng Nam', 'Quảng Ngãi', 'Bình Định',
    'Phú Yên', 'Khánh Hòa', 'Ninh Thuận', 'Bình Thuận',
    # Tay Nguyen (5)
    'Kon Tum', 'Gia Lai', 'Đắk Lắk', 'Đắk Nông', 'Lâm Đồng',
    # Dong Nam Bo (6)
    'Bình Phước', 'Tây Ninh', 'Bình Dương', 'Đồng Nai', 'Bà Rịa - Vũng Tàu',
    'TP. Hồ Chí Minh',
    # Dong bang song Cuu Long (13)
    'Long An', 'Tiền Giang', 'Bến Tre', 'Trà Vinh', 'Vĩnh Long', 'Đồng Tháp',
    'An Giang', 'Kiên Giang', 'Cần Thơ', 'Hậu Giang', 'Sóc Trăng', 'Bạc Liêu',
    'Cà Mau',
]
assert len(PROVINCES_63) == 63

REGIONS_6 = [
    'Đồng bằng sông Hồng',
    'Trung du và miền núi phía Bắc',
    'Bắc Trung Bộ và Duyên hải miền Trung',
    'Tây Nguyên',
    'Đông Nam Bộ',
    'Đồng bằng sông Cửu Long',
]

# ten tieng Anh cua vung (dung khi dong so lieu mang nhan EN)
REGIONS_EN = {
    'red river delta': 'Đồng bằng sông Hồng',
    'northern midlands and mountain areas': 'Trung du và miền núi phía Bắc',
    'northern midlands and mountainous areas': 'Trung du và miền núi phía Bắc',
    'north central and central coastal areas': 'Bắc Trung Bộ và Duyên hải miền Trung',
    'north central area and central coastal area': 'Bắc Trung Bộ và Duyên hải miền Trung',
    'central highlands': 'Tây Nguyên',
    'south east': 'Đông Nam Bộ',
    'mekong river delta': 'Đồng bằng sông Cửu Long',
}

WHOLE_COUNTRY = ('cả nước', 'whole country', 'ca nuoc')
# "TONG SO - TOTAL" chi duoc coi la dong ca nuoc khi VIET HOA toan bo
# (phan biet voi header cot "Tong so/Total" viet thuong)
WHOLE_COUNTRY_UPPER = ('tổng số', 'total', 'tongso')


def _strip_accents(s: str) -> str:
    s = s.replace('đ', 'd').replace('Đ', 'D').replace('ƣ', 'u').replace('Ƣ', 'U')
    s = unicodedata.normalize('NFD', s)
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


def norm_key(s: str) -> str:
    """Khoa so khop: bo dau, bo ky tu ngan cach, thuong hoa."""
    s = _strip_accents(s.strip().lower())
    s = re.sub(r'\b(tp|thanh pho)\.?\s*', '', s)
    return re.sub(r'[^a-z0-9]', '', s)


_PROV_KEYS = {norm_key(p): p for p in PROVINCES_63}
# Bi danh ten tinh (giu he 63 tinh truoc sap nhap):
# - Tu NGTK2024 "Thua Thien - Hue" viet gon la "Hue"/"Thanh pho Hue"
_PROV_ALIASES = {
    'hue': 'Thừa Thiên - Huế',
    'thanhphohue': 'Thừa Thiên - Huế',
    'tphue': 'Thừa Thiên - Huế',
}
_PROV_KEYS.update({norm_key(k): v for k, v in _PROV_ALIASES.items()})
_REG_KEYS = {norm_key(r): r for r in REGIONS_6}
_REG_KEYS.update({norm_key(k): v for k, v in REGIONS_EN.items()})


_WHOLE_KEYS = {norm_key(w) for w in WHOLE_COUNTRY}


def _match_one(k: str):
    if k in _WHOLE_KEYS:
        return 'CA_NUOC', 'CẢ NƯỚC'
    if k in _PROV_KEYS:
        return 'TINH', _PROV_KEYS[k]
    if k in _REG_KEYS:
        return 'VUNG', _REG_KEYS[k]
    return None


_TONGSO_KEYS = {norm_key(w) for w in WHOLE_COUNTRY_UPPER}


def classify_label(label: str):
    """Tra ve (loai, ten_chuan): CA_NUOC / VUNG / TINH / KHAC.
    Nhan co the song ngu tren 1 dong ('CẢ NƯỚC - WHOLE COUNTRY',
    'Đồng bằng sông Hồng - Red River Delta') -> thu ca cac doan cat theo ' - '.
    """
    k = norm_key(label)
    if not k:
        return 'KHAC', label
    # "TONG SO - TOTAL" viet hoa toan bo = dong ca nuoc
    if label == label.upper():
        for tk in _TONGSO_KEYS:
            if k == tk or k.startswith(tk) or k.endswith(tk):
                return 'CA_NUOC', 'CẢ NƯỚC'
    m = _match_one(k)
    if m:
        return m
    parts = [p.strip() for p in label.split(' - ') if p.strip()]
    # thu cac tien to ghep dai truoc ('Thừa Thiên - Huế' trong nhan dai hon)
    for end in range(len(parts), 0, -1):
        m = _match_one(norm_key(' - '.join(parts[:end])))
        if m:
            return m
    # nhan vung/ca nuoc song ngu bi xuong dong nhieu lan ("Trung du va mien nui
    # phia Bac Northern midlands and mountain areas"): so TIEN TO voi khoa vung
    for key, name in _REG_KEYS.items():
        if len(key) >= 8 and k.startswith(key):
            return 'VUNG', name
    for wk in _WHOLE_KEYS:
        if len(wk) >= 6 and k.startswith(wk):
            return 'CA_NUOC', 'CẢ NƯỚC'
    return 'KHAC', label


def slugify(s: str, maxlen=60) -> str:
    s = _strip_accents(s.lower())
    s = re.sub(r'[^a-z0-9]+', '-', s).strip('-')
    return s[:maxlen].rstrip('-')
