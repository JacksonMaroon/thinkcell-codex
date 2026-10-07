"""Guarded standard COM file-moniker serialization, without Office automation.

FileMoniker layout: MS-OSHARED section 2.3.7.8. New serialization comes from
Windows CreateFileMoniker/OleSaveToStream, including its legitimate ANSI path.
"""
from __future__ import annotations
import ctypes
import os
import struct

FILE_CLSID = bytes.fromhex('0303000000000000c000000000000046')
COMPOSITE_CLSID = bytes.fromhex('0903000000000000c000000000000046')
ITEM_CLSID = bytes.fromhex('0403000000000000c000000000000046')


def file_extent(payload: bytes, offset: int = 0) -> tuple[int, str]:
    """Validate one absolute file moniker and return its end and long path."""
    if payload[offset:offset + 16] != FILE_CLSID:
        raise RuntimeError('unsupported file-moniker CLSID')
    start = offset + 16
    if len(payload) < start + 6:
        raise RuntimeError('truncated file moniker')
    anti, length = struct.unpack_from('<HI', payload, start)
    if anti or not 1 <= length <= 32767:
        raise RuntimeError('unsupported relative or oversized file moniker')
    ansi_start = start + 6
    ansi_end = ansi_start + length
    if len(payload) < ansi_end + 28 or payload[ansi_end - 1:ansi_end] != b'\0':
        raise RuntimeError('truncated ANSI moniker path')
    tail = payload[ansi_end:ansi_end + 24]
    if tail != b'\xff\xff\xad\xde' + b'\0' * 20:
        raise RuntimeError('unsupported file-moniker serialization header')
    size = struct.unpack_from('<I', payload, ansi_end + 24)[0]
    end = ansi_end + 28
    path = payload[ansi_start:ansi_end - 1].decode('cp1252')
    if size:
        if size < 6 or len(payload) < end + size:
            raise RuntimeError('truncated Unicode file-moniker extent')
        count, key = struct.unpack_from('<IH', payload, end)
        if count + 6 != size or count % 2 or key != 3:
            raise RuntimeError('invalid Unicode file-moniker size or key')
        path = payload[end + 6:end + size].decode('utf-16-le', 'strict')
        end += size
    if not path or '\0' in path:
        raise RuntimeError('invalid moniker path')
    return end, path


def composite_file(payload: bytes) -> tuple[int, int, str]:
    if payload[:16] != COMPOSITE_CLSID or payload[16:20] != struct.pack('<I', 2):
        raise RuntimeError('unsupported composite-moniker profile')
    end, path = file_extent(payload, 20)
    if payload[end:end + 16] != ITEM_CLSID:
        raise RuntimeError('file moniker must be followed by the original item moniker')
    return 20, end, path


def item_name(payload: bytes, offset: int = 0) -> str:
    """Read the standard two length-prefixed item-moniker strings.

    Windows may append UTF16 text to the NUL-terminated ANSI string inside
    either length-prefixed field. Prefer that authoritative Unicode text.
    """
    if payload[offset:offset + 16] != ITEM_CLSID:
        raise RuntimeError('unsupported item-moniker CLSID')
    cursor = offset + 16
    strings = []
    for _ in range(2):
        if cursor + 4 > len(payload):
            raise RuntimeError('truncated item-moniker string length')
        length = struct.unpack_from('<I', payload, cursor)[0]
        cursor += 4
        if not length or cursor + length > len(payload):
            raise RuntimeError('truncated item-moniker string')
        value = payload[cursor:cursor + length]
        cursor += length
        ansi, separator, unicode = value.partition(b'\0')
        if not separator:
            raise RuntimeError('item-moniker ANSI string is unterminated')
        text = unicode.decode('utf-16-le', 'strict') if unicode else ansi.decode('cp1252')
        if not text or '\0' in text:
            raise RuntimeError('invalid item-moniker name')
        strings.append(text)
    if strings[0] != '!' or cursor != len(payload):
        raise RuntimeError('unsupported item delimiter or trailing moniker bytes')
    return strings[1]


def serialize_file(path: str) -> bytes:
    raw = _serialize(path)
    end, readback = file_extent(raw)
    if end != len(raw) or readback.casefold() != path.casefold():
        raise RuntimeError('Windows file-moniker path readback failed')
    return raw


def serialize_item(name: str) -> bytes:
    if not name or '\0' in name:
        raise ValueError('item moniker requires a nonempty name without NUL')
    raw = _serialize(name, item=True)
    if item_name(raw) != name:
        raise RuntimeError('Windows item-moniker serialization failed')
    return raw


def serialize_composite(path: str, item: str) -> bytes:
    return COMPOSITE_CLSID + struct.pack('<I', 2) + serialize_file(path) + serialize_item(item)


def _serialize(name: str, item: bool = False) -> bytes:
    """Use documented Windows COM exports; never open Excel or PowerPoint."""
    if os.name != 'nt':
        raise RuntimeError('arbitrary-length native moniker serialization requires Windows')
    from ctypes import wintypes
    ole = ctypes.OleDLL('ole32')
    kernel = ctypes.WinDLL('kernel32')
    ole.CoInitializeEx.argtypes = [ctypes.c_void_p, wintypes.DWORD]
    ole.CoInitializeEx.restype = ctypes.c_long
    initialized = ole.CoInitializeEx(None, 2)
    if initialized < 0 and initialized != -2147417850:  # RPC_E_CHANGED_MODE
        raise RuntimeError(f'COM initialization failed: {initialized}')
    moniker, stream, memory = ctypes.c_void_p(), ctypes.c_void_p(), ctypes.c_void_p()
    locked = False
    def checked(value):
        if value < 0:
            raise RuntimeError(f'Windows moniker serialization failed: {value}')
    def release(pointer):
        if pointer.value:
            table = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
            ctypes.WINFUNCTYPE(wintypes.ULONG, ctypes.c_void_p)(table[2])(pointer)
    try:
        if item:
            ole.CreateItemMoniker.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_void_p)]
            checked(ole.CreateItemMoniker('!', name, ctypes.byref(moniker)))
        else:
            ole.CreateFileMoniker.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_void_p)]
            checked(ole.CreateFileMoniker(name, ctypes.byref(moniker)))
        ole.CreateStreamOnHGlobal.argtypes = [ctypes.c_void_p, wintypes.BOOL, ctypes.POINTER(ctypes.c_void_p)]
        checked(ole.CreateStreamOnHGlobal(None, True, ctypes.byref(stream)))
        ole.OleSaveToStream.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        checked(ole.OleSaveToStream(moniker, stream))
        # IStream::Seek is slot 5 after IUnknown and ISequentialStream. Obtain
        # logical length instead of assuming GlobalSize equals stream size.
        table = ctypes.cast(stream, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        seek = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.c_longlong,
                                 wintypes.DWORD, ctypes.POINTER(ctypes.c_ulonglong))(table[5])
        length = ctypes.c_ulonglong()
        checked(seek(stream, 0, 2, ctypes.byref(length)))
        ole.GetHGlobalFromStream.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)]
        checked(ole.GetHGlobalFromStream(stream, ctypes.byref(memory)))
        kernel.GlobalLock.argtypes = [ctypes.c_void_p]; kernel.GlobalLock.restype = ctypes.c_void_p
        kernel.GlobalSize.argtypes = [ctypes.c_void_p]; kernel.GlobalSize.restype = ctypes.c_size_t
        address = kernel.GlobalLock(memory)
        if not address:
            raise RuntimeError('unable to read serialized native moniker')
        locked = True
        if not 0 < length.value <= kernel.GlobalSize(memory):
            raise RuntimeError('invalid serialized moniker stream size')
        return ctypes.string_at(address, length.value)
    finally:
        if locked:
            kernel.GlobalUnlock.argtypes = [ctypes.c_void_p]
            kernel.GlobalUnlock(memory)
        release(stream); release(moniker)
        if initialized >= 0:
            ole.CoUninitialize()
