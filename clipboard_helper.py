"""Windows clipboard: Unicode text plus full-resolution image, no auto-send."""
import ctypes
from ctypes import wintypes
import io
import os
import time
from PIL import Image


def image_formats(path):
    """Prepare original dimensions; preview resizing never affects clipboard."""
    with Image.open(path) as source:
        rgb = source.convert("RGB")
        bmp = io.BytesIO()
        rgb.save(bmp, format="BMP")
        png = io.BytesIO()
        source.save(png, format="PNG")
    return png.getvalue(), bmp.getvalue()[14:]


def copy_content(hwnd, text=None, image_path=None):
    if os.name != "nt":
        raise RuntimeError("Salin gambar tersedia pada Windows.")
    user = ctypes.WinDLL("user32", use_last_error=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    user.OpenClipboard.argtypes = [wintypes.HWND]
    user.OpenClipboard.restype = wintypes.BOOL
    user.CloseClipboard.argtypes = []
    user.CloseClipboard.restype = wintypes.BOOL
    user.EmptyClipboard.argtypes = []
    user.EmptyClipboard.restype = wintypes.BOOL
    user.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user.SetClipboardData.restype = wintypes.HANDLE
    user.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    user.RegisterClipboardFormatW.restype = wintypes.UINT
    kernel.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel.GlobalLock.restype = ctypes.c_void_p
    kernel.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel.GlobalUnlock.restype = wintypes.BOOL
    kernel.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel.GlobalFree.restype = wintypes.HGLOBAL
    formats = []
    if image_path is not None:
        png, dib = image_formats(image_path)
        png_id = user.RegisterClipboardFormatW("PNG")
        if not png_id:
            raise ctypes.WinError(ctypes.get_last_error())
        formats.extend([(png_id, png), (8, dib)])  # CF_DIB
    if text is not None:
        formats.append((13, (text.replace("\x00", "") + "\x00").encode("utf-16-le")))
    if not formats:
        raise ValueError("Tidak ada isi untuk disalin.")
    allocations = []
    try:
        # Allocate before emptying the clipboard. Ownership transfers only on success.
        for fmt, payload in formats:
            handle = kernel.GlobalAlloc(0x0002, len(payload))
            if not handle:
                raise ctypes.WinError(ctypes.get_last_error())
            allocations.append([fmt, handle])
            ptr = kernel.GlobalLock(handle)
            if not ptr:
                raise ctypes.WinError(ctypes.get_last_error())
            try:
                ctypes.memmove(ptr, payload, len(payload))
            finally:
                kernel.GlobalUnlock(handle)
        for attempt in range(10):
            if user.OpenClipboard(hwnd):
                break
            if attempt == 9:
                raise RuntimeError("Clipboard sedang digunakan. Coba tombol salin kembali.")
            time.sleep(0.05)
        try:
            if not user.EmptyClipboard():
                raise ctypes.WinError(ctypes.get_last_error())
            for item in allocations:
                if not user.SetClipboardData(item[0], item[1]):
                    raise RuntimeError("Salinan belum lengkap. Coba salin kembali.")
                item[1] = None
        finally:
            user.CloseClipboard()
    finally:
        for _, handle in allocations:
            if handle:
                kernel.GlobalFree(handle)
