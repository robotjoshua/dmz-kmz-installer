"""Test the full install flow without GUI.

Usage:
    python test_install.py <path_to_kmz> [drone_code]

Example:
    python test_install.py mission.kmz 89
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from waypoint_map_installer import (
    get_mtp_devices, list_missions, replace_mission, verify_mission_on_device
)

if len(sys.argv) < 2:
    print("Usage: python test_install.py <path_to_kmz> [drone_code]")
    print("  drone_code defaults to 89 (Mini 4 Pro)")
    sys.exit(1)

KMZ_FILE = sys.argv[1]
TARGET_DRONE = sys.argv[2] if len(sys.argv) > 2 else "89"

# Step 1: Find device
devices = get_mtp_devices()
print(f"Devices: {len(devices)}")
for d in devices:
    print(f"  {d['friendly_name']} ({d['device_id'][:40]}...)")

if not devices:
    print("No devices found. Connect a DJI RC via USB.")
    sys.exit(1)

device_id = devices[0]["device_id"]

# Step 2: List missions
missions = list_missions(device_id)
print(f"\nMissions: {len(missions)}")
for m in missions:
    print(f"  {m['folder_name']} (has_kmz={m['has_kmz']})")

if not missions:
    print("No missions found on device.")
    sys.exit(1)

mission = missions[0]["folder_name"]

# Step 3: Verify BEFORE install
print(f"\n=== BEFORE INSTALL ===")
report = verify_mission_on_device(device_id, mission)
print(report)

# Step 4: Install
print(f"\n=== INSTALLING ===")
ok, detail = replace_mission(
    device_id, mission, KMZ_FILE,
    target_drone_code=TARGET_DRONE,
    progress_callback=lambda msg: print(f"  {msg}")
)
print(f"\nResult: ok={ok}")
print(f"Detail: {detail[:500]}")

# Step 5: Verify AFTER install
print(f"\n=== AFTER INSTALL ===")
report = verify_mission_on_device(device_id, mission)
print(report)
