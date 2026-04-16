"""Test MediaDevices.dll from Python via pythonnet."""
import clr
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
clr.AddReference("MediaDevices")
from MediaDevices import MediaDevice

print("MediaDevices loaded!")
devices = MediaDevice.GetDevices()
for d in devices:
    print(f"Device: {d.Description} / {d.FriendlyName}")
    d.Connect()
    print(f"  Connected: {d.IsConnected}")

    # List root directories
    try:
        root_path = "\\"
        dirs = d.GetDirectories(root_path)
        for dd in dirs:
            print(f"  Dir: {dd}")
    except Exception as e:
        print(f"  Error listing root: {e}")

    # Try to navigate to waypoint path
    wp_path = r"\Internal shared storage\Android\data\dji.go.v5\files\waypoint"
    try:
        if d.DirectoryExists(wp_path):
            print(f"\n  Waypoint path exists!")
            mission_dirs = d.GetDirectories(wp_path)
            for md in mission_dirs:
                print(f"    Mission dir: {md}")
                try:
                    files = d.GetFiles(md)
                    for f in files:
                        print(f"      File: {f}")
                except Exception as e2:
                    print(f"      Error: {e2}")
        else:
            print(f"  Waypoint path NOT found, trying alternatives...")
            storage_dirs = d.GetDirectories(root_path)
            for sd in storage_dirs:
                print(f"    Storage: {sd}")
    except Exception as e:
        print(f"  Error: {e}")

    d.Disconnect()
    print("  Disconnected")
