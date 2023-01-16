$target_folder = Read-Host "Enter the folder path:"
cd $target_folder
$success = $true
for ($i = 7; $i -le 9; $i++) {
    if (Test-Path "Drake0$i") {
        cd "Drake0$i\FT"
        if (Test-Path "DRAKE0$i.exe") {
            $exe_path = (Get-Item "DRAKE0$i.exe").FullName
            $exe_name = "DRAKE0$i"
            Write-Host "Creating shortcut for $exe_name on Desktop"
            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$exe_name.lnk")
            $Shortcut.TargetPath = $exe_path
            $Shortcut.Save()
            Write-Host "Creating shortcut for $exe_name in Start Menu"
            $Shortcut = $WshShell.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$exe_name.lnk")
            $Shortcut.TargetPath = $exe_path
            $Shortcut.Save()
        } else {
            $success = $false
        }
        cd ..\..
    } else {
        $success = $false
    }
}
for ($i = 10; $i -le 22; $i++) {
    if (Test-Path "Drake$i") {
        cd "Drake$i\FT"
        if (Test-Path "DRAKE$i.exe") {
            $exe_path = (Get-Item "DRAKE$i.exe").FullName
            $exe_name = "DRAKE$i"
            Write-Host "Creating shortcut for $exe_name on Desktop"
            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$exe_name.lnk")
            $Shortcut.TargetPath = $exe_path
            $Shortcut.Save()
            Write-Host "Creating shortcut for $exe_name in Start Menu"
            $Shortcut = $WshShell.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$exe_name.lnk")
            $Shortcut.TargetPath = $exe_path
            $Shortcut.Save()
        } else {
            $success = $false
        }
        cd ..\..
    } else {
        $success = $false
    }
}

if ($success) {
    Write-Host "All shortcuts created successfully" -ForegroundColor Green
    timeout /t 5
} else {
    Write-Host "Some shortcuts were not created successfully" -ForegroundColor Red
    cmd /c 'pause'
}