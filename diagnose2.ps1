$shell = New-Object -ComObject Shell.Application
$mypc = $shell.Namespace(17)

foreach ($item in $mypc.Items()) {
    if ($item.Name -like '*DJI*') {
        $devNs = $shell.Namespace($item.Path)
        foreach ($vol in $devNs.Items()) {
            $cur = $vol.GetFolder
            $path = @("Android", "data", "dji.go.v5", "files", "waypoint")
            foreach ($part in $path) {
                $next = $null
                foreach ($f in $cur.Items()) { if ($f.Name -eq $part) { $next = $f; break } }
                $cur = $next.GetFolder
            }

            # For each mission, copy the mission file BACK to PC so we can inspect it
            $outDir = Join-Path $PSScriptRoot "device_dump"
            New-Item -ItemType Directory -Force -Path $outDir | Out-Null

            foreach ($f in $cur.Items()) {
                if (-not $f.IsFolder) { continue }
                if ($f.Name -eq 'capability' -or $f.Name -eq 'map_preview') { continue }

                $missionName = $f.Name
                Write-Output "MISSION: $missionName"

                $sub = $f.GetFolder
                if ($sub) {
                    foreach ($sf in $sub.Items()) {
                        Write-Output "  FILE: $($sf.Name) | IsFolder=$($sf.IsFolder)"
                        if (-not $sf.IsFolder) {
                            # Copy file back to PC
                            $missionDir = Join-Path $outDir $missionName
                            New-Item -ItemType Directory -Force -Path $missionDir | Out-Null
                            $destFolder = $shell.Namespace($missionDir)
                            $destFolder.CopyHere($sf, 0x0614)
                            Start-Sleep -Seconds 3

                            # Check what we got
                            $copied = Get-ChildItem $missionDir
                            foreach ($c in $copied) {
                                Write-Output "  COPIED: $($c.Name) ($($c.Length) bytes)"
                            }
                        }
                    }
                }
                Write-Output ""
            }
        }
    }
}
