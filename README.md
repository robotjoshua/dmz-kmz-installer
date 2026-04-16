# DPM KMZ Installer

A Windows desktop app for transferring KMZ waypoint missions to DJI RC 2 / RC Pro controllers over USB (MTP).

Built for pilots who need to push custom waypoint missions directly to their DJI controller without going through DJI Fly's cloud workflow.

---

## Features

- Detects connected DJI RC controllers via MTP
- Lists existing waypoint missions on the device
- Normalizes KMZ structure (flattens subfolder layout DJI requires)
- Auto-patches `droneEnumValue` to match your target drone model
- Uploads preview images so the mission shows correctly in DJI Fly
- Supports DJI Mini 3 Pro, Mini 4 Pro, Air 3, Air 3S, Mavic 3, Mavic 3 Pro, Mavic 3 Enterprise, Matrice 350

## Requirements

- Windows 10/11
- Python 3.10+ (for running from source)
- .NET Framework 4.7.2+ (for MediaDevices.dll)
- DJI RC 2 or RC Pro connected via USB in MTP mode

## Running from Source

```bash
pip install -r requirements.txt
python waypoint_map_installer.py
```

## Building the Executable

```bash
pip install -r requirements.txt
python build.py
# Output: dist/DPMKMZInstaller.exe
```

## Usage

1. Connect your DJI RC via USB — set it to **File Transfer (MTP)** mode if prompted
2. Launch the app and click **Refresh** to detect the controller
3. Click **Browse KMZ...** and select your `.kmz` waypoint file
4. Select the target drone model (auto-patches the KMZ if needed)
5. Click **List Missions** to see existing missions on the device
6. Select the mission slot you want to replace
7. Click **Install to RC**
8. Restart the DJI RC once when all installs are done

## Dependencies

| Package | License | Purpose |
|---------|---------|---------|
| [pythonnet](https://github.com/pythonnet/pythonnet) | MIT | .NET interop for MediaDevices |
| [Pillow](https://python-pillow.org/) | HPND | Preview image generation |
| [MediaDevices](https://github.com/Bassman2/MediaDevices) | MIT | MTP file operations via .NET |

`MediaDevices.dll` is bundled in this repo for convenience. It is MIT licensed — see LICENSE for full attribution.

## License

MIT — see [LICENSE](LICENSE)
