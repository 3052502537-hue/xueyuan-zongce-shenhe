# -*- coding: utf-8 -*-
"""
ui_render · Pillow 渲染（抗锯齿圆角 + 柔和投影）
============================================================================
把"画圆角矩形"这件 tk 不擅长的事统一用 Pillow 离线渲染好，再以 PhotoImage
贴回 tk 控件。所有 PhotoImage 都按 key 缓存，避免同一形状反复渲染。

公开 API：
    card_photo(w, h, radius, fill, border, shadow=True)  -> 卡片底图
    btn_photo(w, h, radius, fill, border)               -> 按钮底图
    pill_photo(w, h, fill)                              -> 胶囊底图
    logo_photo(size=34)                                 -> 品牌标记（主色 + 白色对勾）

⚠ 必须在有 Tk 上下文（root 已创建）后调用，因为 ImageTk.PhotoImage 依赖 Tcl。
"""
try:
    from PIL import Image, ImageDraw, ImageFilter, ImageTk
    HAVE_PIL = True
except Exception:   # Pillow 缺失时降级：控件会改用 tk 自带矩形（略糙但不崩）
    HAVE_PIL = False
    Image = ImageDraw = ImageFilter = ImageTk = None


from .ui_design import (
    C_ACCENT, SHADOW_M, SHADOW_BLUR, SHADOW_OFF, SHADOW_ALPHA,
)


# 抗锯齿超采样倍数：先把图放大 3x 再下采样回目标尺寸
_SS_CARD = 2
_SS_BTN  = 3

# 缓存上限：超过即按 FIFO 丢弃最早的 key
_CACHE_LIMIT = 500
_PHOTO_CACHE = {}
_PHOTO_ORDER = []


def _scale_method():
    """Pillow 1.1+ 起 Resampling 是枚举；旧版直接挂在 Image 类上。"""
    if not HAVE_PIL:
        return None
    return getattr(Image, 'Resampling', Image).LANCZOS


def _cache(key, maker):
    """按 key 缓存 PhotoImage，FIFO 淘汰。"""
    ph = _PHOTO_CACHE.get(key)
    if ph is not None:
        return ph
    ph = ImageTk.PhotoImage(maker())
    _PHOTO_CACHE[key] = ph
    _PHOTO_ORDER.append(key)
    if len(_PHOTO_ORDER) > _CACHE_LIMIT:
        old = _PHOTO_ORDER.pop(0)
        _PHOTO_CACHE.pop(old, None)
    return ph


def _hex2rgb(h):
    h = str(h).lstrip('#')
    if len(h) == 3:
        h = ''.join(c * 2 for c in h)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _rounded_img(w, h, radius, fill, border, ss, shadow=False,
                 m=0, blur=0, off=0, alpha=0):
    """在超采样画布上画一个带可选阴影的圆角矩形，再 resize 到目标尺寸。

    ss: 超采样倍数（值越大越平滑，但越慢；通常 2~4）
    """
    w, h = max(1, int(w)), max(1, int(h))
    if w < 4 or h < 4:
        return Image.new('RGB', (w, h), (255, 255, 255))
    W, H = max(1, w * ss), max(1, h * ss)
    if shadow and m > 0:
        m = int(min(m, max(1, (min(w, h) - 3) // 2)))
    img = Image.new('RGBA', (W, H), (255, 255, 255, 255))

    # —— 阴影层 ——
    if shadow and m > 0:
        sx0, sy0 = m * ss, (m + off) * ss
        sx1, sy1 = (w - m) * ss, (h - m + off) * ss
        if sx1 <= sx0:
            sx1 = sx0 + ss
        if sy1 <= sy0:
            sy1 = sy0 + ss
        shl = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(shl)
        d.rounded_rectangle([sx0, sy0, sx1, sy1], radius=radius * ss,
                            fill=(0, 0, 0, alpha))
        shl = shl.filter(ImageFilter.GaussianBlur(max(1, blur * ss)))
        img = Image.alpha_composite(img, shl)

    # —— 主体圆角矩形 ——
    cx0, cy0 = m * ss, m * ss
    cx1, cy1 = (w - m) * ss - 1, (h - m) * ss - 1
    if cx1 <= cx0:
        cx1 = cx0 + ss
    if cy1 <= cy0:
        cy1 = cy0 + ss
    card = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(card)
    d.rounded_rectangle([cx0, cy0, cx1, cy1], radius=radius * ss,
                        fill=_hex2rgb(fill) + (255,),
                        outline=_hex2rgb(border) + (255,), width=max(1, ss // 2))
    img = Image.alpha_composite(img, card)
    img = img.convert('RGB')
    if ss != 1:
        img = img.resize((w, h), _scale_method())
    return img


def _safe_size(w, h):
    return max(2, int(w)), max(2, int(h))


def card_photo(w, h, radius, fill, border, shadow=True):
    """卡片底图：圆角 + 浅阴影。"""
    w, h = _safe_size(w, h)
    r = int(radius)
    m = SHADOW_M if shadow else 0

    def maker():
        return _rounded_img(w, h, r, fill, border, _SS_CARD, shadow,
                            m=m, blur=SHADOW_BLUR, off=SHADOW_OFF,
                            alpha=SHADOW_ALPHA)
    return _cache(('card', w, h, r, fill, border, shadow), maker)


def btn_photo(w, h, radius, fill, border):
    """按钮底图：圆角（无阴影）。"""
    w, h = _safe_size(w, h)
    r = int(radius)

    def maker():
        return _rounded_img(w, h, r, fill, border, _SS_BTN)
    return _cache(('btn', w, h, r, fill, border), maker)


def pill_photo(w, h, fill):
    """胶囊底图：左右半圆。"""
    w, h = _safe_size(w, h)

    def maker():
        return _rounded_img(w, h, h, fill, fill, _SS_BTN)
    return _cache(('pill', w, h, fill), maker)


def logo_photo(size=34):
    """品牌标记：主色圆角方块 + 白色对勾（矢量绘制，缩放不糊）。"""
    s = int(size)

    def maker():
        ss = 4
        W = s * ss
        img = Image.new('RGBA', (W, W), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([0, 0, W - 1, W - 1], radius=int(W * 0.28),
                            fill=_hex2rgb(C_ACCENT) + (255,))
        lw = max(ss, int(W * 0.11))
        pts = [(W * 0.27, W * 0.53), (W * 0.44, W * 0.70), (W * 0.75, W * 0.33)]
        d.line(pts, fill=(255, 255, 255, 255), width=lw, joint='curve')
        for p in (pts[0], pts[2]):
            d.ellipse([p[0] - lw / 2, p[1] - lw / 2,
                       p[0] + lw / 2, p[1] + lw / 2],
                      fill=(255, 255, 255, 255))
        return img.convert('RGB').resize((s, s), _scale_method())

    return _cache(('logo', s), maker)


def clear_cache():
    """清空所有 PhotoImage 缓存（仅供测试 / 主题切换时使用）。"""
    _PHOTO_CACHE.clear()
    _PHOTO_ORDER.clear()