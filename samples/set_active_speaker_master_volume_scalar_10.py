# アクティブなスピーカーのマスター音量を10%に設定する。

from powccoreaudio.devicepropsinstore import DevicePropertiesReadOnlyInPropertyStore
from powccoreaudio.mmdevice import MMDeviceEnumerator

enumerator = MMDeviceEnumerator.create()
audio_device = enumerator.get_speaker()

audio = audio_device.activate_audioendpointvolume()
audio.master_volume_level_scalar = 0.1

# 確認用
props = DevicePropertiesReadOnlyInPropertyStore(audio_device.propstore_read)
print(f"{props.friendlyname}: {audio.master_volume_level_scalar * 100:.2f}%")
