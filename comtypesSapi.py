# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "comtypes",
# ]
# ///
from ctypes import (
    POINTER,
    Structure,
    c_longlong,
    c_ubyte,
    c_ulong,
    c_ulonglong,
    c_wchar_p,
    _Pointer,
)
from ctypes.wintypes import DWORD
from typing import TYPE_CHECKING
import comtypes.client
from comtypes import COMObject, IUnknown, hresult, GUID, COMMETHOD, HRESULT


class FILETIME(Structure):
    _fields_ = (
        ("dwLowDateTime", DWORD),
        ("dwHighDateTime", DWORD),
    )


WSTRING = c_wchar_p


class _LARGE_INTEGER(Structure):
    _fields_ = [
        ("QuadPart", c_longlong),
    ]


class _ULARGE_INTEGER(Structure):
    _fields_ = [
        ("QuadPart", c_ulonglong),
    ]


class tagSTATSTG(Structure):
    _fields_ = [
        ("pwcsName", WSTRING),
        ("type", c_ulong),
        ("cbSize", _ULARGE_INTEGER),
        ("mtime", FILETIME),
        ("ctime", FILETIME),
        ("atime", FILETIME),
        ("grfMode", c_ulong),
        ("grfLocksSupported", c_ulong),
        ("clsid", GUID),
        ("grfStateBits", c_ulong),
        ("reserved", c_ulong),
    ]


class ISequentialStream(IUnknown):
    _iid_ = GUID("{0C733A30-2A1C-11CE-ADE5-00AA0044773D}")
    _idlflags_ = []
    _methods_ = [
        COMMETHOD(
            [],
            HRESULT,
            "RemoteRead",
            (["out"], POINTER(c_ubyte), "pv"),
            (["in"], c_ulong, "cb"),
            (["out"], POINTER(c_ulong), "pcbRead"),
        ),
        COMMETHOD(
            [],
            HRESULT,
            "RemoteWrite",
            (["in"], POINTER(c_ubyte), "pv"),
            (["in"], c_ulong, "cb"),
            (["out"], POINTER(c_ulong), "pcbWritten"),
        ),
    ]


class IStream(ISequentialStream):
    _iid_ = GUID("{0000000C-0000-0000-C000-000000000046}")
    _idlflags_ = []


IStream._methods_ = [
    COMMETHOD(
        [],
        HRESULT,
        "RemoteSeek",
        (["in"], _LARGE_INTEGER, "dlibMove"),
        (["in"], c_ulong, "dwOrigin"),
        (["out"], POINTER(_ULARGE_INTEGER), "plibNewPosition"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "SetSize",
        (["in"], _ULARGE_INTEGER, "libNewSize"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "RemoteCopyTo",
        (["in"], POINTER(IStream), "pstm"),
        (["in"], _ULARGE_INTEGER, "cb"),
        (["out"], POINTER(_ULARGE_INTEGER), "pcbRead"),
        (["out"], POINTER(_ULARGE_INTEGER), "pcbWritten"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "Commit",
        (["in"], c_ulong, "grfCommitFlags"),
    ),
    COMMETHOD([], HRESULT, "Revert"),
    COMMETHOD(
        [],
        HRESULT,
        "LockRegion",
        (["in"], _ULARGE_INTEGER, "libOffset"),
        (["in"], _ULARGE_INTEGER, "cb"),
        (["in"], c_ulong, "dwLockType"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "UnlockRegion",
        (["in"], _ULARGE_INTEGER, "libOffset"),
        (["in"], _ULARGE_INTEGER, "cb"),
        (["in"], c_ulong, "dwLockType"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "Stat",
        (["out"], POINTER(tagSTATSTG), "pstatstg"),
        (["in"], c_ulong, "grfStatFlag"),
    ),
    COMMETHOD(
        [],
        HRESULT,
        "Clone",
        (["out"], POINTER(POINTER(IStream)), "ppstm"),
    ),
]


if TYPE_CHECKING:
    LP_c_ubyte = _Pointer[c_ubyte]
    LP_c_ulong = _Pointer[c_ulong]
    LP__ULARGE_INTEGER = _Pointer[_ULARGE_INTEGER]
else:
    LP_c_ubyte = POINTER(c_ubyte)
    LP_c_ulong = POINTER(c_ulong)
    LP__ULARGE_INTEGER = POINTER(_ULARGE_INTEGER)


class AudioStream(COMObject):
    _com_interfaces_ = [IStream]

    def ISequentialStream_RemoteWrite(
        self,
        this: int,
        pv: LP_c_ubyte,
        cb: int,
        pcbWritten: LP_c_ulong,
    ) -> int:
        return hresult.S_OK

    def IStream_RemoteSeek(
        self,
        this: int,
        dlibMove: _LARGE_INTEGER,
        dwOrigin: int,
        plibNewPosition: LP__ULARGE_INTEGER,
    ) -> int:
        return hresult.S_OK

    def IStream_Commit(self, grfCommitFlags: int):
        pass


def main():
    tts = comtypes.client.CreateObject("SAPI.SPVoice")
    audioStream = AudioStream()
    customStream = comtypes.client.CreateObject("SAPI.SpCustomStream")
    customStream.BaseStream = audioStream
    tts.AudioOutputStream = customStream


if __name__ == "__main__":
    main()
