# powccoreaudio
PythonからWindowsのCoreAudio COM機能を使いやすくするパッケージです。comtypes、powcpropsys、powdeviceinfoに依存します。 

**マスター音量を10%に設定する**

```python
from powccoreaudio.mmdevice import MMDeviceEnumerator
from powccoreaudio.devicepropsinstore import DevicePropertiesReadOnlyInPropertyStore

enumerator = MMDeviceEnumerator.create()
audio_device = enumerator.get_speaker()

audio = audio_device.activate_audioendpointvolume()
audio.master_volume_level_scalar = 0.1

# 確認用
props = DevicePropertiesReadOnlyInPropertyStore(audio_device.propstore_read)
print(f"{props.friendlyname}: {audio.master_volume_level_scalar * 100:.2f}%")
```

**全スピーカーのミュート**

```python
from powccoreaudio.mmdevice import MMDeviceEnumerator

device_enum = MMDeviceEnumerator.create()
for device in device_enum.speakers:
    volume_ctrl = device.activate_audioendpointvolume()
    volume_ctrl.mute = True
```
