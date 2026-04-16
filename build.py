"""
Build script — packages DPM KMZ Installer into a standalone .exe
Run:  python build.py
"""
import PyInstaller.__main__
import os

HERE = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    os.path.join(HERE, "waypoint_map_installer.py"),
    "--name", "DPMKMZInstaller",
    "--onefile",
    "--windowed",                       # no console window
    "--icon", os.path.join(HERE, "assets", "waypointmap.ico"),
    # Bundle the DLL and assets into the exe
    "--add-data", f"{os.path.join(HERE, 'MediaDevices.dll')};.",
    "--add-data", f"{os.path.join(HERE, 'assets')};assets",
    # pythonnet hidden imports
    "--hidden-import", "clr_loader",
    "--hidden-import", "clr_loader.ffi",
    "--hidden-import", "clr_loader.util.find",
    "--hidden-import", "clr_loader.util.clr_error",
    "--collect-all", "clr_loader",
    "--collect-all", "pythonnet",
    # Work dirs
    "--distpath", os.path.join(HERE, "dist"),
    "--workpath", os.path.join(HERE, "build"),
    "--specpath", HERE,
    # Overwrite previous build
    "--noconfirm",
])

print("\nBuild complete! Your exe is at: dist/DPMKMZInstaller.exe")
