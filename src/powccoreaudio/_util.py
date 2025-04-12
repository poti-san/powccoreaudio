from powcpropsys.propstore import PropertyKey
from powdeviceinfo.devprop import DevicePropertyKey


def _devpkey_to_pkey(devpkey: DevicePropertyKey) -> PropertyKey:
    return PropertyKey(devpkey)
