from ctypes import POINTER, Structure, byref, c_float, c_int32, c_uint32, c_void_p
from dataclasses import dataclass
from enum import IntFlag
from typing import Any

from comtypes import GUID, STDMETHOD, IUnknown
from powc.core import ComResult, cr, query_interface


class IAudioEndpointVolume(IUnknown):
    _iid_ = GUID("{5CDF2C82-841E-4546-9722-0CF74078229A}")
    _methods_ = [
        # TODO: IAudioEndpointVolumeCallback
        STDMETHOD(c_int32, "RegisterControlChangeNotify", (POINTER(IUnknown),)),
        # TODO: IAudioEndpointVolumeCallback
        STDMETHOD(c_int32, "UnregisterControlChangeNotify", (POINTER(IUnknown),)),
        STDMETHOD(c_int32, "GetChannelCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "SetMasterVolumeLevel", (c_float, POINTER(GUID))),
        STDMETHOD(c_int32, "SetMasterVolumeLevelScalar", (c_float, POINTER(GUID))),
        STDMETHOD(c_int32, "GetMasterVolumeLevel", (POINTER(c_float),)),
        STDMETHOD(c_int32, "GetMasterVolumeLevelScalar", (POINTER(c_float),)),
        STDMETHOD(c_int32, "SetChannelVolumeLevel", (c_uint32, c_float, POINTER(GUID))),
        STDMETHOD(c_int32, "SetChannelVolumeLevelScalar", (c_uint32, c_float, POINTER(GUID))),
        STDMETHOD(c_int32, "GetChannelVolumeLevel", (c_uint32, POINTER(c_float))),
        STDMETHOD(c_int32, "GetChannelVolumeLevelScalar", (c_uint32, POINTER(c_float))),
        STDMETHOD(c_int32, "SetMute", (c_int32, POINTER(GUID))),
        STDMETHOD(c_int32, "GetMute", (POINTER(c_int32),)),
        STDMETHOD(c_int32, "GetVolumeStepInfo", (POINTER(c_uint32), POINTER(c_uint32))),
        STDMETHOD(c_int32, "VolumeStepUp", (POINTER(GUID),)),
        STDMETHOD(c_int32, "VolumeStepDown", (POINTER(GUID),)),
        STDMETHOD(c_int32, "QueryHardwareSupport", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetVolumeRange", (POINTER(c_float), POINTER(c_float), POINTER(c_float))),
    ]

    __slots__ = ()


class AudioVolumeNotificationData(Structure):
    """AUDIO_VOLUME_NOTIFICATION_DATA"""

    __slots__ = ()
    _fields_ = (
        ("event_context_guid", GUID),
        ("muted", c_int32),
        ("master_volume", c_float),
        ("channels", c_uint32),
        ("channel_volumes0", c_float),
    )


class ENDPOINT_HARDWARE_SUPPORT(IntFlag):
    """ENDPOINT_HARDWARE_SUPPORT_*定数"""

    VOLUME = 0x00000001
    MUTE = 0x00000002
    METER = 0x00000004


class IAudioEndpointVolumeCallback(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{657804FA-D6AD-4496-8A60-352752AF4F89}")
    _methods_ = [
        STDMETHOD(c_int32, "OnNotify", (POINTER(AudioVolumeNotificationData),)),
    ]


class IAudioEndpointVolumeEx(IAudioEndpointVolume):
    __slots__ = ()
    _iid_ = GUID("{66E11784-F695-4F28-A505-A7080081A78F}")
    _methods_ = [
        STDMETHOD(c_int32, "GetVolumeRangeChannel", (c_uint32, POINTER(c_float), POINTER(c_float), POINTER(c_float))),
    ]


class IAudioMeterInformation(IUnknown):
    __slots__ = ()
    _iid_ = GUID("{C02216F6-8C67-4B5B-9D00-D008E73E0064}")
    _methods_ = [
        STDMETHOD(c_int32, "GetPeakValue", (POINTER(c_float),)),
        STDMETHOD(c_int32, "GetMeteringChannelCount", (POINTER(c_uint32),)),
        STDMETHOD(c_int32, "GetChannelsPeakValues", (c_uint32, POINTER(c_float))),
        STDMETHOD(c_int32, "QueryHardwareSupport", (POINTER(c_uint32),)),
    ]


class AudioEndpointVolume:
    """オーディオエンドポイント。IAudioEndpointVolumeインターフェイスのラッパーです。"""

    __slots__ = ("__o", "__eventcontext_guid")
    __o: Any  # POINTER(IAudioEndpointVolume)
    __eventcontext_guid: GUID | None

    def __init__(self, o: Any, eventcontext_guid: GUID | None = None) -> None:
        self.__o = query_interface(o, IAudioEndpointVolume)
        self.__eventcontext_guid = eventcontext_guid

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def eventcontext_guid(self) -> GUID | None:
        return self.__eventcontext_guid

    @eventcontext_guid.setter
    def eventcontext_guid(self, guid: GUID | None):
        self.__eventcontext_guid = guid

    #         # TODO: IAudioEndpointVolumeCallback
    #         STDMETHOD(c_int32, "RegisterControlChangeNotify", (POINTER(IUnknown),)),
    #         # TODO: IAudioEndpointVolumeCallback
    #         STDMETHOD(c_int32, "UnregisterControlChangeNotify", (POINTER(IUnknown),)),

    @property
    def channel_count_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetChannelCount(x), x.value)

    @property
    def channel_count(self) -> int:
        return self.channel_count_nothrow.value

    @property
    def master_volume_level_nothrow(self) -> ComResult[float]:
        x = c_float()
        return cr(self.__o.GetMasterVolumeLevel(byref(x)), x.value)

    @property
    def master_volume_level_scalar_nothrow(self) -> ComResult[float]:
        x = c_float()
        return cr(self.__o.GetMasterVolumeLevelScalar(byref(x)), x.value)

    def set_master_volume_level_nothrow(self, value: float) -> ComResult[None]:
        return cr(self.__o.SetMasterVolumeLevel(value, self.__eventcontext_guid), None)

    def set_master_volume_level_scalar_nothrow(self, value: float) -> ComResult[None]:
        return cr(self.__o.SetMasterVolumeLevelScalar(value, self.__eventcontext_guid), None)

    @property
    def master_volume_level(self) -> float:
        x = c_float()
        return cr(self.__o.GetMasterVolumeLevel(byref(x)), x.value).value

    @property
    def master_volume_level_scalar(self) -> float:
        x = c_float()
        return cr(self.__o.GetMasterVolumeLevelScalar(byref(x)), x.value).value

    @master_volume_level.setter
    def master_volume_level(self, value: float) -> None:
        return self.set_master_volume_level_nothrow(value).value

    @master_volume_level_scalar.setter
    def master_volume_level_scalar(self, value: float) -> None:
        return self.set_master_volume_level_scalar_nothrow(value).value

    def set_channel_volume_level_nothrow(self, channel: int, value: float) -> ComResult[None]:
        return cr(self.__o.SetChannelVolumeLevel(channel, value, self.__eventcontext_guid), None)

    def set_channel_volume_level_scalar_nothrow(self, channel: int, value: float) -> ComResult[None]:
        return cr(self.__o.SetChannelVolumeLevelScalar(channel, value, self.__eventcontext_guid), None)

    def set_channel_volume_level(self, channel: int, value: float) -> None:
        return self.set_channel_volume_level_nothrow(channel, value).value

    def set_channel_volume_level_scalar(self, channel: int, value: float) -> None:
        return self.set_channel_volume_level_scalar_nothrow(channel, value).value

    def get_channel_volume_level_nothrow(self, channel: int) -> ComResult[float]:
        x = c_float()
        return cr(self.__o.GetChannelVolumeLevel(channel, byref(x)), x.value)

    def get_channel_volume_level_scalar_nothrow(self, channel: int) -> ComResult[float]:
        x = c_float()
        return cr(self.__o.GetChannelVolumeLevelScalar(channel, byref(x)), x.value)

    @property
    def mute_nothrow(self) -> ComResult[bool]:
        x = c_int32()
        return cr(self.__o.GetMute(byref(x)), x.value != 0)

    @property
    def mute(self) -> bool:
        return self.mute_nothrow.value

    def set_mute_nothrow(self, f: bool) -> ComResult[None]:
        return cr(self.__o.SetMute(1 if f else 0, self.__eventcontext_guid), None)

    @mute.setter
    def mute(self, f: bool) -> None:
        return self.set_mute_nothrow(f).value

    @dataclass(frozen=True)
    class VolumeStepInfo:
        step: int
        step_count: int

    @property
    def volume_step_info_nothrow(self) -> ComResult[VolumeStepInfo]:
        x1 = c_uint32()
        x2 = c_uint32()
        return cr(
            self.__o.GetVolumeStepInfo(byref(x1), byref(x2)), AudioEndpointVolume.VolumeStepInfo(x1.value, x2.value)
        )

    @property
    def volume_step_info(self) -> VolumeStepInfo:
        return self.volume_step_info_nothrow.value

    def volume_stepup_nothrow(self) -> ComResult[None]:
        return cr(self.__o.VolumeStepUp(self.__eventcontext_guid), None)

    def volume_stepup(self) -> None:
        return self.volume_stepup_nothrow().value

    def volume_stepdown_nothrow(self) -> ComResult[None]:
        return cr(self.__o.VolumeStepDown(self.__eventcontext_guid), None)

    def volume_stepdown(self) -> None:
        return self.volume_stepdown_nothrow().value

    @property
    def hardware_support_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.QueryHardwareSupport(byref(x)), x.value)

    @property
    def hardware_support(self) -> int:
        return self.hardware_support_nothrow.value

    @dataclass(frozen=True)
    class VolumeRange:
        min_db: float
        max_db: float
        increment_db: float

    @property
    def volume_range_nothrow(self) -> ComResult[VolumeRange]:
        x1 = c_float()
        x2 = c_float()
        x3 = c_float()
        return cr(
            self.__o.GetVolumeRange(byref(x1), byref(x2), byref(x3)),
            AudioEndpointVolume.VolumeRange(x1.value, x2.value, x3.value),
        )

    @property
    def volume_range(self) -> VolumeRange:
        return self.volume_range_nothrow.value


class AudioMeterInformation:
    """聴覚メーター情報。IAudioMeterInformationインターフェイスのラッパーです。

    作成にはMMDeviceを使用します。"""

    __o: Any  # POINTER(IAudioMeterInformation)

    __slots__ = ("__o",)

    def __init__(self, o: Any) -> None:
        self.__o = query_interface(o, IAudioMeterInformation)

    @property
    def wrapped_obj(self) -> c_void_p:
        return self.__o

    @property
    def peakvalue_nothrow(self) -> ComResult[float]:
        x = c_float()
        return cr(self.__o.GetPeakValue(byref(x)), x.value)

    @property
    def peakvalue(self) -> float:
        return self.peakvalue_nothrow.value

    @property
    def meteringchannelcount_nothrow(self) -> ComResult[int]:
        x = c_uint32()
        return cr(self.__o.GetMeteringChannelCount(byref(x)), x.value)

    @property
    def meteringchannelcount(self) -> int:
        return self.meteringchannelcount_nothrow.value

    @property
    def channelspeakvalues_nothrow(self) -> ComResult[tuple[float]]:
        c = self.meteringchannelcount_nothrow
        if not c:
            return cr(c.hr, tuple())
        x = (c_float * c.value_unchecked)()
        return cr(self.__o.GetChannelsPeakValues(c.value_unchecked, x), tuple(x))

    @property
    def channelspeakvalues(self) -> tuple[float]:
        return self.channelspeakvalues_nothrow.value

    @property
    def hardwaresupport_nothrow(self) -> ComResult[ENDPOINT_HARDWARE_SUPPORT]:
        x = c_uint32()
        return cr(self.__o.QueryHardwareSupport(byref(x)), ENDPOINT_HARDWARE_SUPPORT(x.value))

    @property
    def hardwaresupport(self) -> ENDPOINT_HARDWARE_SUPPORT:
        return self.hardwaresupport_nothrow.value
