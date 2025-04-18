"""スピーカーのメーター情報"""

from powccoreaudio.mmdevice import MMDeviceEnumerator

device_enum = MMDeviceEnumerator.create()
for device in device_enum.speakers:
    audiometerinfo = device.activate_audiometerinfo()
    pass
