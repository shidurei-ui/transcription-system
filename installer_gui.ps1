Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$REPO_URL    = "https://github.com/shidurei-ui/transcription-system.git"
$INSTALL_DIR = Join-Path $env:USERPROFILE "transcription-system"

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Transcription System Installer" Height="620" Width="520"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="#1a1a2e">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
    </Window.Resources>
    <Grid Margin="30">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" HorizontalAlignment="Center" Margin="0,0,0,10">
            <TextBlock Text="Transcription System" FontSize="28" FontWeight="Bold"
                       HorizontalAlignment="Center" Foreground="#e94560"/>
            <TextBlock Text="Automatic lecture transcription" FontSize="14"
                       HorizontalAlignment="Center" Foreground="#888" Margin="0,5,0,0"/>
        </StackPanel>

        <Border Grid.Row="1" Background="#16213e" CornerRadius="10" Padding="15" Margin="0,5,0,10">
            <StackPanel>
                <TextBlock Text="Auto-installs everything:" FontSize="13" Foreground="#ccc" Margin="0,0,0,5"/>
                <TextBlock Text="  V  Python" FontSize="13" Foreground="#7fdbca"/>
                <TextBlock Text="  V  ffmpeg" FontSize="13" Foreground="#7fdbca"/>
                <TextBlock Text="  V  Required packages" FontSize="13" Foreground="#7fdbca"/>
                <TextBlock Text="  V  Auto-start with Windows" FontSize="13" Foreground="#7fdbca"/>
                <TextBlock Text="" FontSize="6"/>
                <TextBlock Text="What you need:" FontSize="13" FontWeight="Bold" Foreground="White" Margin="0,5,0,3"/>
                <TextBlock Text="  1. Gemini API key (free)" FontSize="13" Foreground="#ffd460"/>
                <TextBlock Text="  2. Load Chrome extension (once)" FontSize="13" Foreground="#ffd460"/>
            </StackPanel>
        </Border>

        <Border Grid.Row="2" Background="#16213e" CornerRadius="10" Padding="15" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="Gemini API Key" FontSize="14" FontWeight="Bold" Margin="0,0,0,8"/>
                <TextBlock FontSize="11" Foreground="#52b4ff" Margin="0,0,0,8" TextWrapping="Wrap"
                           Cursor="Hand" Name="linkText" TextDecorations="Underline"
                           Text="Click here to get a free key: aistudio.google.com/app/apikey"/>
                <TextBox Name="apiKeyBox" Height="35" FontSize="14" Padding="8,5"
                         Background="#0f3460" Foreground="White" BorderBrush="#e94560"
                         BorderThickness="1" FlowDirection="LeftToRight"
                         VerticalContentAlignment="Center"/>
            </StackPanel>
        </Border>

        <Button Name="installBtn" Grid.Row="3" Height="45"
                FontSize="16" FontWeight="Bold" Cursor="Hand"
                Foreground="White" Background="#e94560" BorderThickness="0"
                Margin="0,0,0,10" Content="Install">
            <Button.Resources>
                <Style TargetType="Border">
                    <Setter Property="CornerRadius" Value="10"/>
                </Style>
            </Button.Resources>
        </Button>

        <Border Grid.Row="4" Background="#0f0f23" CornerRadius="10" Padding="10" Margin="0,0,0,10">
            <ScrollViewer Name="logScroll" VerticalScrollBarVisibility="Auto">
                <TextBlock Name="logBox" TextWrapping="Wrap" FontSize="12"
                           FontFamily="Consolas" Foreground="#7fdbca" FlowDirection="LeftToRight"/>
            </ScrollViewer>
        </Border>

        <ProgressBar Name="progressBar" Grid.Row="5" Height="6"
                     Background="#16213e" Foreground="#e94560" BorderThickness="0"
                     Margin="0,0,0,10" Visibility="Collapsed"/>

        <TextBlock Name="statusText" Grid.Row="6" Text="Waiting..."
                   FontSize="12" Foreground="#888" HorizontalAlignment="Center"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$apiKeyBox   = $window.FindName("apiKeyBox")
$installBtn  = $window.FindName("installBtn")
$logBox      = $window.FindName("logBox")
$logScroll   = $window.FindName("logScroll")
$progressBar = $window.FindName("progressBar")
$statusText  = $window.FindName("statusText")
$linkText    = $window.FindName("linkText")

function Add-Log($msg, $color) {
    if (-not $color) { $color = "#7fdbca" }
    $window.Dispatcher.Invoke([Action]{
        $run = New-Object System.Windows.Documents.Run($msg)
        $run.Foreground = $color
        $logBox.Inlines.Add($run)
        $logBox.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
        $logScroll.ScrollToEnd()
    })
}

function Set-Prog($pct, $status) {
    $window.Dispatcher.Invoke([Action]{
        $progressBar.Value = $pct
        $statusText.Text = $status
    })
}

function Refresh-Path {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","User")
}

$linkText.Add_MouseLeftButtonDown({
    Start-Process "https://aistudio.google.com/app/apikey"
})

$script:installDone = $false

$installBtn.Add_Click({
    if ($script:installDone) {
        $extDir = Join-Path $INSTALL_DIR "extension"
        try { Start-Process "chrome" "chrome://extensions" } catch {
            try { Start-Process "msedge" "chrome://extensions" } catch {}
        }
        $nl = [Environment]::NewLine
        $msg = "1. Enable Developer mode (top right)" + $nl + "2. Click Load unpacked" + $nl + "3. Select folder:" + $nl + $extDir
        [System.Windows.MessageBox]::Show($msg, "Load Chrome Extension", "OK", "Information")
        return
    }

    $apiKey = $apiKeyBox.Text.Trim()
    if ($apiKey.Length -lt 10) {
        [System.Windows.MessageBox]::Show("Please enter a valid Gemini API key", "Error", "OK", "Warning")
        return
    }

    $installBtn.IsEnabled = $false
    $apiKeyBox.IsEnabled  = $false
    $progressBar.Visibility = "Visible"

    $sd  = $INSTALL_DIR
    $ru  = $REPO_URL
    $key = $apiKey

    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("window", $window)
    $runspace.SessionStateProxy.SetVariable("logBox", $logBox)
    $runspace.SessionStateProxy.SetVariable("logScroll", $logScroll)
    $runspace.SessionStateProxy.SetVariable("progressBar", $progressBar)
    $runspace.SessionStateProxy.SetVariable("statusText", $statusText)
    $runspace.SessionStateProxy.SetVariable("installBtn", $installBtn)
    $runspace.SessionStateProxy.SetVariable("sd", $sd)
    $runspace.SessionStateProxy.SetVariable("ru", $ru)
    $runspace.SessionStateProxy.SetVariable("key", $key)

    $ps = [PowerShell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({

        function Add-Log($msg, $color) {
            if (-not $color) { $color = "#7fdbca" }
            $window.Dispatcher.Invoke([Action]{
                $run = New-Object System.Windows.Documents.Run($msg)
                $run.Foreground = $color
                $logBox.Inlines.Add($run)
                $logBox.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
                $logScroll.ScrollToEnd()
            })
        }
        function Set-Prog($pct, $status) {
            $window.Dispatcher.Invoke([Action]{
                $progressBar.Value = $pct
                $statusText.Text = $status
            })
        }
        function Refresh-Path {
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                        [System.Environment]::GetEnvironmentVariable("Path","User")
        }

        try {
            # 1. Git
            Add-Log "Checking Git..." "#ffd460"
            Set-Prog 5 "Checking Git..."
            $gitOk = $false
            try { $null = & git --version 2>&1; $gitOk = $true } catch {}

            if (-not $gitOk) {
                Add-Log "Downloading and installing Git..." "#ffd460"
                Set-Prog 8 "Installing Git..."
                $gitInst = Join-Path $env:TEMP "git_installer.exe"
                Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/download/v2.47.1.windows.2/Git-2.47.1.2-64-bit.exe" -OutFile $gitInst -UseBasicParsing
                Start-Process -FilePath $gitInst -ArgumentList "/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS" -Wait
                Remove-Item $gitInst -Force -ErrorAction SilentlyContinue
                Refresh-Path
                Add-Log "[OK] Git installed" "#7fdbca"
            } else {
                Add-Log "[OK] Git found" "#7fdbca"
            }

            # 2. Clone
            Set-Prog 15 "Downloading system..."
            if (Test-Path $sd) {
                Add-Log "Folder exists, updating..." "#ffd460"
                Push-Location $sd
                try { & git pull --quiet 2>&1 } catch {}
                Pop-Location
            } else {
                Add-Log "Cloning repository..." "#ffd460"
                & git clone $ru $sd 2>&1
            }
            Add-Log "[OK] System files downloaded" "#7fdbca"

            # 3. Python
            Set-Prog 25 "Checking Python..."
            Add-Log "Checking Python..." "#ffd460"
            $python = $null
            foreach ($cmd in @("python","python3","py")) {
                try {
                    $ver = & $cmd --version 2>&1
                    if ($ver -match "Python 3") { $python = $cmd; break }
                } catch {}
            }
            if (-not $python) {
                Add-Log "Downloading Python 3.11..." "#ffd460"
                Set-Prog 30 "Installing Python..."
                $pyInst = Join-Path $env:TEMP "python_installer.exe"
                Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe" -OutFile $pyInst -UseBasicParsing
                Start-Process -FilePath $pyInst -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1" -Wait
                Remove-Item $pyInst -Force -ErrorAction SilentlyContinue
                Refresh-Path
                foreach ($cmd in @("python","python3","py")) {
                    try {
                        $ver = & $cmd --version 2>&1
                        if ($ver -match "Python 3") { $python = $cmd; break }
                    } catch {}
                }
                if (-not $python) { Add-Log "[FAIL] Python not installed" "#ff6b6b"; return }
                Add-Log "[OK] Python installed" "#7fdbca"
            } else {
                Add-Log "[OK] Python found" "#7fdbca"
            }

            # 4. pip packages
            Set-Prog 45 "Installing packages..."
            Add-Log "Installing Python packages..." "#ffd460"
            & $python -m pip install --upgrade pip --quiet 2>&1
            & $python -m pip install fastapi uvicorn yt-dlp "google-genai" python-dotenv pyaudiowpatch psutil pywin32 opencv-python --quiet 2>&1
            Add-Log "[OK] Packages installed" "#7fdbca"

            # 5. ffmpeg
            Set-Prog 60 "Checking ffmpeg..."
            Add-Log "Checking ffmpeg..." "#ffd460"
            $ffmpegOk = $false
            try { $null = & ffmpeg -version 2>&1; $ffmpegOk = $true } catch {}
            $localFf = Join-Path $sd "ffmpeg.exe"
            if ((-not $ffmpegOk) -and (Test-Path $localFf)) { $ffmpegOk = $true }

            if (-not $ffmpegOk) {
                Add-Log "Downloading ffmpeg (~60MB)..." "#ffd460"
                Set-Prog 65 "Downloading ffmpeg..."
                $ffZip = Join-Path $env:TEMP "ffmpeg.zip"
                $ffTmp = Join-Path $env:TEMP "ffmpeg_extracted"
                Invoke-WebRequest -Uri "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl-shared.zip" -OutFile $ffZip -UseBasicParsing
                Expand-Archive -Path $ffZip -DestinationPath $ffTmp -Force
                $ffExe = Get-ChildItem -Path $ffTmp -Recurse -Filter "ffmpeg.exe" | Select-Object -First 1
                Copy-Item $ffExe.FullName -Destination $localFf -Force
                Get-ChildItem (Join-Path $ffExe.DirectoryName "*.dll") | ForEach-Object {
                    Copy-Item $_.FullName -Destination $sd -Force
                }
                Remove-Item $ffZip -Force -ErrorAction SilentlyContinue
                Remove-Item $ffTmp -Recurse -Force -ErrorAction SilentlyContinue
                Add-Log "[OK] ffmpeg installed" "#7fdbca"
            } else {
                Add-Log "[OK] ffmpeg found" "#7fdbca"
            }

            # 6. PATH
            Set-Prog 80 "Setting PATH..."
            $userPath = [System.Environment]::GetEnvironmentVariable("Path","User")
            if ($userPath -notlike "*$sd*") {
                [System.Environment]::SetEnvironmentVariable("Path", "$sd;$userPath", "User")
            }

            # 7. API Key
            Set-Prog 85 "Saving API key..."
            Add-Log "Saving API key..." "#ffd460"
            "GEMINI_API_KEY=$key" | Set-Content (Join-Path $sd ".env") -Encoding UTF8
            Add-Log "[OK] API key saved" "#7fdbca"

            # 8. Unblock
            Get-ChildItem -Path $sd -Recurse -File |
                ForEach-Object { try { Unblock-File $_.FullName } catch {} }

            # 9. Startup
            Set-Prog 90 "Setting auto-start..."
            Add-Log "Setting auto-start with Windows..." "#ffd460"
            $vbsPath = Join-Path $sd "start_server_silent.vbs"
            if (Test-Path $vbsPath) {
                $startupDir = [System.Environment]::GetFolderPath("Startup")
                $WshShell = New-Object -ComObject WScript.Shell
                $sc = $WshShell.CreateShortcut((Join-Path $startupDir "transcription-server.lnk"))
                $sc.TargetPath = "wscript.exe"
                $sc.Arguments = """$vbsPath"""
                $sc.WorkingDirectory = $sd
                $sc.Save()
                Add-Log "[OK] Auto-start configured" "#7fdbca"
            }

            # 10. Start server
            Set-Prog 95 "Starting server..."
            Add-Log "Starting server..." "#ffd460"
            if (Test-Path $vbsPath) {
                Start-Process "wscript.exe" """$vbsPath"""
                Add-Log "[OK] Server running in background" "#7fdbca"
            }

            # Done
            Set-Prog 100 "Installation complete!"
            Add-Log "" "#7fdbca"
            Add-Log "======================================" "#e94560"
            Add-Log "  Installation complete!" "#e94560"
            Add-Log "======================================" "#e94560"
            Add-Log "" "#7fdbca"
            Add-Log "One last step: Load the Chrome extension" "#ffd460"
            Add-Log "Click the button below" "#ffd460"

            $window.Dispatcher.Invoke([Action]{
                $script:installDone = $true
                $installBtn.IsEnabled = $true
                $installBtn.Content = "Open Chrome to load extension"
            })

        } catch {
            Add-Log ("[FAIL] Error: " + $_.Exception.Message) "#ff6b6b"
            Set-Prog 0 "Installation failed"
            $window.Dispatcher.Invoke([Action]{
                $installBtn.IsEnabled = $true
            })
        }
    }) | Out-Null

    $ps.BeginInvoke() | Out-Null
})

$window.ShowDialog() | Out-Null
