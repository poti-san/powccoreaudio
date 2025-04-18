from abc import abstractmethod
from ctypes import (
    POINTER,
    Structure,
    _Pointer,
    byref,
    c_int32,
    c_uint32,
    c_void_p,
    c_wchar_p,
)
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
from powc.core import (
    ComResult,
    IUnknownWrapper,
    check_hresult,
    cotaskmem,
    cr,
    query_interface,
)
from powc.stream import StorageMode
from powcpropsys.propkey import PropertyKey
from powcpropsys.propstore import IPropertyStore, PropertyStore
from powcpropsys.propvariant import PropVariant

from powccoreaudio.devicepropsinstore import DevicePropertiesReadOnlyInPropertyStore
from powccoreaudio.endpointvolume import AudioEndpointVolume, IAudioEndpointVolume

from .endpointvolume import AudioMeterInformation, IAudioMeterInformation


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

    def activate_nothrow[TWrapper: IUnknownWrapper](
        self, iid: GUID, wrapper_type: type[TWrapper], params: PropVariant | None = None, clsctx: int = CLSCTX_ALL
    ) -> ComResult[TWrapper]:

        p = POINTER(IUnknown)()
        return cr(self.__o.Activate(byref(iid), clsctx, params, byref(p)), wrapper_type(p))

    def activate[TWrapper: IUnknownWrapper](
        self, iid: GUID, wrapper_type: type[TWrapper], params: PropVariant | None = None, clsctx: int = CLSCTX_ALL
    ) -> TWrapper:
        return self.activate_nothrow(iid, wrapper_type, params, clsctx).value

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
        return self.activate_nothrow(IAudioEndpointVolume._iid_, AudioEndpointVolume)

    def activate_audioendpointvolume(self) -> AudioEndpointVolume:
        return self.activate_audioendpointvolume_nothrow().value

    def activate_audiosystemeffectspropstore_nothrow(
        self, propstore_context: GUID
    ) -> "ComResult[AudioSystemEffectsPropertyStore]":
        pv = PropVariant.init_clsid(propstore_context)
        return self.activate_nothrow(IAudioSystemEffectsPropertyStore._iid_, AudioSystemEffectsPropertyStore, pv)

    def activate_audiosystemeffectspropstore(self, propstore_context: GUID) -> "AudioSystemEffectsPropertyStore":
        return self.activate_audiosystemeffectspropstore_nothrow(propstore_context).value

    def activate_audiometerinfo_nothrow(self) -> "ComResult[AudioMeterInformation]":
        return self.activate_nothrow(IAudioMeterInformation._iid_, AudioMeterInformation)

    def activate_audiometerinfo(self) -> "AudioMeterInformation":
        return self.activate_audiometerinfo_nothrow().value

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


# TODO クラス分け
PKEY_AudioEndpoint_FormFactor = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 0
)
PKEY_AudioEndpoint_ControlPanelPageProvider = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 1
)
PKEY_AudioEndpoint_Association = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 2
)
PKEY_AudioEndpoint_PhysicalSpeakers = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 3
)
PKEY_AudioEndpoint_GUID = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 4
)
PKEY_AudioEndpoint_Disable_SysFx = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 5
)
ENDPOINT_SYSFX_ENABLED = 0x00000000
ENDPOINT_SYSFX_DISABLED = 0x00000001
PKEY_AudioEndpoint_FullRangeSpeakers = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 6
)
PKEY_AudioEndpoint_Supports_EventDriven_Mode = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 7
)
PKEY_AudioEndpoint_JackSubType = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 8
)
PKEY_AudioEndpoint_Default_VolumeInDb = PropertyKey.from_define(
    0x1DA5D803, 0xD492, 0x4EDD, 0x8C, 0x23, 0xE0, 0xC0, 0xFF, 0xEE, 0x7F, 0x0E, 9
)
PKEY_AudioEngine_DeviceFormat = PropertyKey.from_define(
    0xF19F064D, 0x82C, 0x4E27, 0xBC, 0x73, 0x68, 0x82, 0xA1, 0xBB, 0x8E, 0x4C, 0
)
PKEY_AudioEngine_OEMFormat = PropertyKey.from_define(
    0xE4870E26, 0x3CC5, 0x4CD2, 0xBA, 0x46, 0xCA, 0xA, 0x9A, 0x70, 0xED, 0x4, 3
)
PKEY_AudioEndpointLogo_IconEffects = PropertyKey.from_define(
    0xF1AB780D, 0x2010, 0x4ED3, 0xA3, 0xA6, 0x8B, 0x87, 0xF0, 0xF0, 0xC4, 0x76, 0
)
PKEY_AudioEndpointLogo_IconPath = PropertyKey.from_define(
    0xF1AB780D, 0x2010, 0x4ED3, 0xA3, 0xA6, 0x8B, 0x87, 0xF0, 0xF0, 0xC4, 0x76, 1
)
PKEY_AudioEndpointSettings_MenuText = PropertyKey.from_define(
    0x14242002, 0x0320, 0x4DE4, 0x95, 0x55, 0xA7, 0xD8, 0x2B, 0x73, 0xC2, 0x86, 0
)
PKEY_AudioEndpointSettings_LaunchContract = PropertyKey.from_define(
    0x14242002, 0x0320, 0x4DE4, 0x95, 0x55, 0xA7, 0xD8, 0x2B, 0x73, 0xC2, 0x86, 1
)


class DirectXAudioActivationParams(Structure):
    """DIRECTX_AUDIO_ACTIVATION_PARAMS"""

    __slots__ = ()
    _fields_ = (
        ("directx_audio_activation_params", c_uint32),
        ("audio_session_guid", GUID),
        ("audio_stream_flags", c_uint32),
    )


class EndpointFormFactor(IntEnum):
    RemoteNetworkDevice = 0
    Speakers = 1
    LineLevel = 2
    Headphones = 3
    Microphone = 4
    Headset = 5
    Handset = 6
    UnknownDigitalPassthrough = 7
    SPDIF = 8
    DigitalAudioDisplayDevice = 9
    UnknownFormFactor = 10


class IMMDeviceActivator(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{3B0D0EA4-D0A9-4B0E-935B-09516746FAC0}")
    _methods_ = [
        STDMETHOD(
            c_int32, "Activate", (POINTER(GUID), POINTER(IMMDevice), POINTER(PropVariant), POINTER(POINTER(IUnknown)))
        ),
    ]


class IActivateAudioInterfaceAsyncOperation(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{72A22D78-CDE4-431D-B8CC-843A71199B6D}")
    _methods_ = [
        STDMETHOD(c_int32, "GetActivateResult", (POINTER(c_int32), POINTER(POINTER(IUnknown)))),
    ]


class IActivateAudioInterfaceCompletionHandler(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{41D949AB-9862-444A-80F6-C261334DA5EB}")
    _methods_ = [
        STDMETHOD(c_int32, "ActivateCompleted", (POINTER(IActivateAudioInterfaceAsyncOperation),)),
    ]


# STDAPI ActivateAudioInterfaceAsync(
#     _In_ LPCWSTR deviceInterfacePath,
#     _In_ REFIID riid,
#     _In_opt_ PROPVARIANT *activationParams,
#     _In_ IActivateAudioInterfaceCompletionHandler *completionHandler,
#     _COM_Outptr_ IActivateAudioInterfaceAsyncOperation **activationOperation
#     );


class AudioExtensionParams(Structure):
    __slots__ = ()
    _fields_ = (
        ("add_page_param", c_void_p),
        ("endppoint_p", POINTER(IMMDevice)),
        ("pnp_interface_p", POINTER(IMMDevice)),
        ("pnp_devnode_p", POINTER(IMMDevice)),
    )


class AudioSystemEffectsPropStoreType(IntEnum):
    """AUDIO_SYSTEMEFFECTS_PROPERTYSTORE_TYPE"""

    DEFAULT = 0
    USER = 1
    VOLATILE = 2


class IAudioSystemEffectsPropertyChangeNotificationClient(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{20049D40-56D5-400E-A2EF-385599FEED49}")
    _methods_ = [
        STDMETHOD(c_int32, "OnPropertyChanged", (c_int32, PropertyKey)),
    ]


class IAudioSystemEffectsPropertyStore(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{302AE7F9-D7E0-43E4-971B-1F8293613D2A}")
    _methods_ = [
        STDMETHOD(c_int32, "OpenDefaultPropertyStore", (c_int32, POINTER(POINTER(IPropertyStore)))),
        STDMETHOD(c_int32, "OpenUserPropertyStore", (c_int32, POINTER(POINTER(IPropertyStore)))),
        STDMETHOD(c_int32, "OpenVolatilePropertyStore", (c_uint32, POINTER(POINTER(IPropertyStore)))),
        STDMETHOD(c_int32, "ResetUserPropertyStore", ()),
        STDMETHOD(c_int32, "ResetVolatilePropertyStore", ()),
        STDMETHOD(
            c_int32,
            "RegisterPropertyChangeNotification",
            (POINTER(IAudioSystemEffectsPropertyChangeNotificationClient),),
        ),
        STDMETHOD(
            c_int32,
            "UnregisterPropertyChangeNotification",
            (POINTER(IAudioSystemEffectsPropertyChangeNotificationClient),),
        ),
    ]


class AudioSystemEffectsPropertyStore:
    """聴覚システム効果プロパティストア。IAudioSystemEffectsPropertyStoreインターフェイスのラッパーです。

    作成にはMMDeviceを使用します。"""

    __o: Any  # POINTER(IAudioSystemEffectsPropertyStore)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IAudioSystemEffectsPropertyStore)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    def open_defaultpropstore_nothrow(self, stg: StorageMode) -> ComResult[PropertyStore]:
        x = POINTER(IPropertyStore)()
        return cr(self.__o.OpenDefaultPropertyStore(int(stg)), PropertyStore(x))

    def open_defaultpropstore(self, stg: StorageMode) -> PropertyStore:
        return self.open_defaultpropstore_nothrow(stg).value

    def open_userpropstore_nothrow(self, stg: StorageMode) -> ComResult[PropertyStore]:
        x = POINTER(IPropertyStore)()
        return cr(self.__o.OpenUserPropertyStore(int(stg)), PropertyStore(x))

    def open_userpropstore(self, stg: StorageMode) -> PropertyStore:
        return self.open_userpropstore_nothrow(stg).value

    def open_volatilepropstore_nothrow(self, stg: StorageMode) -> ComResult[PropertyStore]:
        x = POINTER(IPropertyStore)()
        return cr(self.__o.OpenVolatilePropertyStore(int(stg)), PropertyStore(x))

    def open_volatilepropstore(self, stg: StorageMode) -> PropertyStore:
        return self.open_volatilepropstore_nothrow(stg).value

    def reset_userpropstore_nothrow(self) -> ComResult[None]:
        return cr(self.__o.ResetUserPropertyStore(), None)

    def reset_userpropstore(self) -> None:
        return self.reset_userpropstore_nothrow().value

    def reset_volatilepropstore_nothrow(self) -> ComResult[None]:
        return cr(self.__o.ResetVolatilePropertyStore(), None)

    def reset_volatilepropstore(self) -> None:
        return self.reset_volatilepropstore_nothrow().value

    if TYPE_CHECKING:
        IAudioSystemEffectsPropertyChangeNotificationClientType = _Pointer[
            IAudioSystemEffectsPropertyChangeNotificationClient
        ]
    else:
        IAudioSystemEffectsPropertyChangeNotificationClientType = POINTER(
            IAudioSystemEffectsPropertyChangeNotificationClient
        )

    def register_propchangenotification_nothrow(
        self, client: IAudioSystemEffectsPropertyChangeNotificationClientType
    ) -> ComResult[None]:
        return cr(self.__o.RegisterPropertyChangeNotification(client), None)

    def register_propchangenotification(self, client: IAudioSystemEffectsPropertyChangeNotificationClientType) -> None:
        return self.register_propchangenotification_nothrow(client).value

    def unregister_propchangenotification_nothrow(
        self, client: IAudioSystemEffectsPropertyChangeNotificationClientType
    ) -> ComResult[None]:
        return cr(self.__o.UnregisterPropertyChangeNotification(client), None)

    def unregister_propchangenotification(
        self, client: IAudioSystemEffectsPropertyChangeNotificationClientType
    ) -> None:
        return self.unregister_propchangenotification_nothrow(client).value
