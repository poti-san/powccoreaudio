"""全スピーカーのミュート"""

from powccoreaudio.mmdevice import MMDeviceEnumerator

device_enum = MMDeviceEnumerator.create()
for device in device_enum.speakers:
    volume_ctrl = device.activate_audioendpointvolume()
    volume_ctrl.mute = True
