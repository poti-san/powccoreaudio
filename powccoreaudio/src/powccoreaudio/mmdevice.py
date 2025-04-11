from abc import abstractmethod
from ctypes import POINTER, _Pointer, byref, c_int32, c_uint32, c_void_p, c_wchar_p
from enum import IntEnum, IntFlag
from inspect import signature as inspect_sig
from typing import (
    TYPE_CHECKING,
    Any,
    Callable,
    Final,
    Iterator,
    OrderedDict,
    TextIO,
    overload,
    override,
)

from comtypes import CLSCTX_ALL, GUID, STDMETHOD, CoCreateInstance, COMObject, IUnknown
from comtypes.hresult import S_OK
from powc.core import ComResult, check_hresult, cotaskmem, cr, query_interface
from powc.stream import StorageMode
from powccoreaudio.endpointvolume import AudioEndpointVolume, IAudioEndpointVolume
from powcdeviceprop.devicepropsinstore import DevicePropertiesReadOnlyInPropertyStore
from powcpropsys.propkey import PropertyKey
from powcpropsys.propstore import IPropertyStore, PropertyStore
from powcpropsys.propvariant import PropVariant


class DataFlow(IntEnum):
    """EDataFlow"""

    Render = 0
    Capture = 1
    All = 2


class Role(IntEnum):
    """ERole"""

    Console = 0
    Multimedia = 1
    Communications = 2


class DeviceState(IntFlag):
    ACTIVE = 0x00000001
    DISABLED = 0x00000002
    NOTPRESENT = 0x00000004
    UNPLUGGED = 0x00000008


class IMMDevice(IUnknown):
    """"""

    _iid_ = GUID("{D666063F-1587-4E43-81F1-B948E807363F}")
    _methods_ = [
        STDMETHOD(c_int32, "Activate", (POINTER(GUID), c_uint32, POINTER(PropVariant), POINTER(POINTER(IUnknown)))),
        STDMETHOD(c_int32, "OpenPropertyStore", (c_uint32, POINTER(POINTER(IPropertyStore)))),
        STDMETHOD(c_int32, "GetId", (POINTER(c_wchar_p),)),
        STDMETHOD(c_int32, "GetState", (POINTER(c_uint32),)),
    ]
    __slots__ = ()


class MMDevice:
    """マルチメディアデバイス。IMMDeviceインターフェイスのラッパーです。"""

    __o: Any  # POINTER(IMMDevice)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IMMDevice)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    if TYPE_CHECKING:

        def activate_nothrow[TIUnknown: IUnknown](
            self, interface_type: type[TIUnknown], params: PropVariant | None = None, clsctx: int = CLSCTX_ALL
        ) -> ComResult[_Pointer[TIUnknown]]: ...

    else:

        def activate_nothrow[TIUnknown: IUnknown](
            self, interface_type: type[TIUnknown], params: PropVariant | None = None, clsctx: int = CLSCTX_ALL
        ):

            p = POINTER(interface_type)()
            return cr(self.__o.Activate(byref(interface_type._iid_), clsctx, params, byref(p)), p)

    def open_propstore_nothrow(self, access: StorageMode) -> ComResult[PropertyStore]:
        p = POINTER(IPropertyStore)()
        return cr(self.__o.OpenPropertyStore(int(access), byref(p)), PropertyStore(p))

    def open_propstore(self, access: StorageMode) -> PropertyStore:
        return self.open_propstore_nothrow(access).value

    @property
    def propstore_read_nothrow(self) -> ComResult[PropertyStore]:
        return self.open_propstore_nothrow(StorageMode.READ)

    @property
    def propstore_read(self) -> PropertyStore:
        return self.open_propstore_nothrow(StorageMode.READ).value

    @property
    def propstore_readwrite_nothrow(self) -> ComResult[PropertyStore]:
        return self.open_propstore_nothrow(StorageMode.READ | StorageMode.WRITE)

    @property
    def propstore_readwrite(self) -> PropertyStore:
        return self.open_propstore_nothrow(StorageMode.READ | StorageMode.WRITE).value

    @property
    def id_nothrow(self) -> ComResult[str]:
        with cotaskmem(c_wchar_p()) as p:
            return cr(self.__o.GetId(byref(p)), p.value or "")

    @property
    def id(self) -> str:
        return self.id_nothrow.value

    @property
    def state_nothrow(self) -> ComResult[DeviceState]:
        x = c_uint32()
        return cr(self.__o.GetState(byref(x)), DeviceState(x.value))

    @property
    def state(self) -> DeviceState:
        return self.state_nothrow.value

    #
    # ユーティリティ
    #

    def activate_audioendpointvolume_nothrow(self) -> ComResult[AudioEndpointVolume]:
        ret = self.activate_nothrow(IAudioEndpointVolume)
        return cr(ret.hr, AudioEndpointVolume(ret.value_unchecked))

    def activate_audioendpointvolume(self) -> AudioEndpointVolume:
        return self.activate_audioendpointvolume_nothrow().value

    @property
    def device_props_readonly_nothrow(self) -> ComResult[DevicePropertiesReadOnlyInPropertyStore]:
        ret = self.propstore_read_nothrow
        return cr(ret.hr, DevicePropertiesReadOnlyInPropertyStore(ret.value_unchecked))

    @property
    def device_props_readonly(self) -> DevicePropertiesReadOnlyInPropertyStore:
        return self.device_props_readonly_nothrow.value


class IMMDeviceCollection(IUnknown):
    """"""

    _iid_ = GUID("{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}")
    _methods_ = [
        STDMETHOD(c_int32, "GetCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "Item", (c_uint32, POINTER(POINTER(IMMDevice)))),
    ]

    __slots__ = ()


class MMDeviceCollection:
    """マルチメディアデバイスコレクション。IMMDeviceCollectionインターフェイスのラッパーです。"""

    __o: Any  # POINTER(IMMDeviceCollection)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IMMDeviceCollection)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def __repr__(self) -> str:
        return "MMDeviceCollection"

    def __str__(self) -> str:
        return "MMDeviceCollection"

    def __len__(self) -> int:
        x = c_uint32()
        check_hresult(self.__o.GetCount(byref(x)))
        return x.value

    def getat(self, index: int) -> MMDevice:
        """
        インデックスを指定して項目を取得します。
        キーが整数固定なので__getitem__より高速です。
        """
        x = POINTER(IMMDevice)()
        check_hresult(self.__o.Item(int(index), byref(x)))
        return MMDevice(x)

    @overload
    def __getitem__(self, key: int) -> MMDevice: ...
    @overload
    def __getitem__(self, key: slice) -> tuple[MMDevice, ...]: ...
    @overload
    def __getitem__(self, key: tuple[slice, ...]) -> tuple[MMDevice, ...]: ...

    def __getitem__(self, key) -> MMDevice | tuple[MMDevice, ...]:
        if isinstance(key, slice):
            return tuple(self.__getitem__(i) for i in range(*key.indices(len(self))))
        if isinstance(key, tuple):
            for subslice in key:
                if not isinstance(subslice, slice):
                    raise TypeError
            return tuple(item for item in (t for t in self.__getitem__(key)))
        else:
            return self.getat(key)

    def __iter__(self) -> Iterator[MMDevice]:
        return (self.getat(i) for i in range(len(self)))

    @property
    def items(self) -> tuple[MMDevice, ...]:
        return tuple(iter(self))


class IMMEndpoint(IUnknown):
    """"""

    _iid_ = GUID("{1BE09788-6894-4089-8586-9A2A6C265AC5}")
    _methods_ = [
        STDMETHOD(c_int32, "GetDataFlow", (POINTER(c_int32),)),
    ]


class MMEndpoint:
    """マルチメディアエンドポイント。IMMEndpointインターフェイスのラッパーです。"""

    __o: Any  # POINTER(IMMEndpoint)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IMMEndpoint)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def dataflow_nothrow(self) -> ComResult[DataFlow]:
        x = c_int32()
        return cr(self.__o.GetDataFlow(byref(x)), DataFlow(x.value))


class IMMNotificationClient(IUnknown):
    """"""

    _iid_ = GUID("{7991EEC9-7E89-4D85-8390-6C703CEC60C0}")
    _methods_ = [
        STDMETHOD(c_int32, "OnDeviceStateChanged", (c_wchar_p, c_uint32)),
        STDMETHOD(c_int32, "OnDeviceAdded", (c_wchar_p,)),
        STDMETHOD(c_int32, "OnDeviceRemoved", (c_wchar_p,)),
        STDMETHOD(c_int32, "OnDefaultDeviceChanged", (c_uint32, c_int32, c_wchar_p)),
        STDMETHOD(c_int32, "OnPropertyValueChanged", (c_wchar_p, PropertyKey)),
    ]


class MMNotificationClientBase:
    @abstractmethod
    def device_state_changed(self, deviceid: str | None, newstate: int) -> int: ...

    @abstractmethod
    def device_added(self, deviceid: str | None) -> int: ...

    @abstractmethod
    def device_removed(self, deviceid: str | None) -> int: ...

    @abstractmethod
    def defaultdevice_changed(self, flow: DataFlow, role: Role, defaultdeviceid: str | None) -> int: ...

    @abstractmethod
    def propvalue_changed(self, deviceid: str | None, key: PropVariant) -> int: ...


class MMNotificationClientForCall(MMNotificationClientBase):

    __f: Callable[[str, OrderedDict[str, Any]], None]

    __slots__ = ("__f",)

    def __init__(self, f: Callable[[str, OrderedDict[str, Any]], None]) -> None:
        self.__f = f
        super().__init__()

    @override
    def device_state_changed(self, deviceid: str | None, newstate: int) -> int:
        self.__f(
            "OnDeviceStateChanged",
            inspect_sig(MMNotificationClientForCall.device_state_changed).bind(self).arguments,
        )
        return S_OK

    @override
    def device_added(self, deviceid: str | None) -> int:
        self.__f(
            "OnDeviceAdded",
            inspect_sig(MMNotificationClientForCall.device_added).bind(self).arguments,
        )
        return S_OK

    @override
    def device_removed(self, deviceid: str | None) -> int:
        self.__f(
            "OnDeviceRemoved",
            inspect_sig(MMNotificationClientForCall.device_removed).bind(self).arguments,
        )
        return S_OK

    @override
    def defaultdevice_changed(self, flow: DataFlow, role: Role, defaultdeviceid: str | None) -> int:
        self.__f(
            "OnDefaultDeviceChanged",
            inspect_sig(MMNotificationClientForCall.defaultdevice_changed).bind(self).arguments,
        )
        return S_OK

    @override
    def propvalue_changed(self, deviceid: str | None, key: PropVariant) -> int:
        self.__f(
            "OnPropertyValueChanged",
            inspect_sig(MMNotificationClientForCall.propvalue_changed).bind(self).arguments,
        )
        return S_OK


class MMNotificationClientForPrint(MMNotificationClientForCall):
    __f: TextIO | None

    __slots__ = ("__f",)

    def __init__(self, f: TextIO | None = None) -> None:
        self.__f = f
        super().__init__(lambda s, dict: print(f"{s}({dict})", file=self.__f))


class MMNotificationClientWrapper(COMObject):
    _com_interfaces_ = [IMMNotificationClient]

    __client: MMNotificationClientBase

    def __init__(self, client: MMNotificationClientBase) -> None:
        self.__client = client

    def IMMNotificationClient_OnDeviceStateChanged(self, this, deviceid: str | None, newstate: int) -> int:
        return self.__client.device_state_changed(deviceid, newstate)

    def IMMNotificationClient_OnDeviceAdded(self, this, deviceid: str | None) -> int:
        return self.__client.device_added(deviceid)

    def IMMNotificationClient_OnDeviceRemoved(self, this, deviceid: str | None) -> int:
        return self.__client.device_removed(deviceid)

    def IMMNotificationClient_OnDefaultDeviceChanged(
        self, this, flow: int, role: int, defaultdeviceid: str | None
    ) -> int:
        return self.__client.defaultdevice_changed(DataFlow(flow), Role(role), defaultdeviceid)

    def IMMNotificationClient_OnPropertyValueChanged(self, this, deviceid: str | None, key: PropVariant) -> int:
        return self.__client.propvalue_changed(deviceid, key)


class IMMDeviceEnumerator(IUnknown):
    """"""

    _iid_ = GUID("{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    _methods_ = [
        STDMETHOD(c_int32, "EnumAudioEndpoints", (c_int32, c_uint32, POINTER(POINTER(IMMDeviceCollection)))),
        STDMETHOD(c_int32, "GetDefaultAudioEndpoint", (c_int32, c_int32, POINTER(POINTER(IMMDevice)))),
        STDMETHOD(c_int32, "GetDevice", (c_wchar_p, POINTER(POINTER(IMMDevice)))),
        STDMETHOD(c_int32, "RegisterEndpointNotificationCallback", (POINTER(IMMNotificationClient),)),
        STDMETHOD(c_int32, "UnregisterEndpointNotificationCallback", (POINTER(IMMNotificationClient),)),
    ]


class MMDeviceEnumerator:
    """マルチメディアデバイスの列挙機能。IMMDeviceEnumeratorインターフェイスのラッパーです。

    Examples:
        >>> device_enum = MMDeviceEnumerator.create()
    """

    CLSID_MMDEVICEENUMERATOR: Final = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")

    __o: Any  # POINTER(IMMDeviceEnumerator)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IMMDeviceEnumerator)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @staticmethod
    def create() -> "MMDeviceEnumerator":
        return MMDeviceEnumerator(CoCreateInstance(MMDeviceEnumerator.CLSID_MMDEVICEENUMERATOR, IMMDeviceEnumerator))

    def get_enum_audioendpoints_nothrow(self, flow: DataFlow, role: Role) -> ComResult[MMDeviceCollection]:
        p = POINTER(IMMDeviceCollection)()
        return cr(self.__o.EnumAudioEndpoints(int(flow), int(role), byref(p)), MMDeviceCollection(p))

    def get_enum_audioendpoints(self, flow: DataFlow, role: Role) -> MMDeviceCollection:
        return self.get_enum_audioendpoints_nothrow(flow, role).value

    def get_default_audioendpoint_nothrow(self, flow: DataFlow, role: Role) -> ComResult[MMDevice]:
        p = POINTER(IMMDevice)()
        return cr(self.__o.GetDefaultAudioEndpoint(int(flow), int(role), byref(p)), MMDevice(p))

    @property
    def speakers(self) -> MMDeviceCollection:
        return self.get_enum_audioendpoints(DataFlow.Render, Role.Multimedia)

    @property
    def microphones(self) -> MMDeviceCollection:
        return self.get_enum_audioendpoints(DataFlow.Capture, Role.Multimedia)

    def get_default_audioendpoint(self, flow: DataFlow, role: Role) -> MMDevice:
        return self.get_default_audioendpoint_nothrow(flow, role).value

    def get_speaker(self) -> MMDevice:
        return self.get_default_audioendpoint(DataFlow.Render, Role.Multimedia)

    def get_microphone(self) -> MMDevice:
        return self.get_default_audioendpoint(DataFlow.Capture, Role.Multimedia)

    def get_device_nothrow(self, id: str) -> ComResult[MMDevice]:
        p = POINTER(IMMDevice)()
        return cr(self.__o.GetDevice(id, byref(p)), MMDevice(p))

    def get_device(self, id: str) -> MMDevice:
        return self.get_device_nothrow(id).value

    def register_endpoint_notification_callback_nothrow(
        self, client: MMNotificationClientBase
    ) -> ComResult[MMNotificationClientWrapper]:

        wrapper = MMNotificationClientWrapper(client)
        return cr(self.__o.RegisterEndpointNotificationCallback(wrapper), wrapper)  # type: ignore

    def register_endpoint_notification_callback(self, client: MMNotificationClientBase) -> MMNotificationClientWrapper:
        return self.register_endpoint_notification_callback_nothrow(client).value

    def unregister_endpoint_notification_callback_nothrow(
        self, wrapper: MMNotificationClientWrapper
    ) -> ComResult[None]:

        return cr(self.__o.UnregisterEndpointNotificationCallback(wrapper), None)

    def unregister_endpoint_notification_callback(self, wrapper: MMNotificationClientWrapper) -> None:
        return self.unregister_endpoint_notification_callback_nothrow(wrapper).value
