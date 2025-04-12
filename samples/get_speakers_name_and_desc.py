"""スピーカーの表示名と概要の列挙"""

from powccoreaudio.mmdevice import MMDeviceEnumerator

device_enum = MMDeviceEnumerator.create()
for device in device_enum.speakers:
    devprops = device.device_props_readonly
    print(f"{devprops.friendlyname} ({device.id}): {devprops.device_desc}")
