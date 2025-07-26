Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "このスクリプトは管理者権限で実行する必要があります。管理者として再度実行してください。" -ForegroundColor Red
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$PSCommandPath`""
    Exit
}

Write-Host "===== Windows dotsfiles 設定スクリプトを開始します =====" -ForegroundColor Green

# 1. WSLのインストールと有効化
Write-Host "WSLを有効化し、インストールを開始します..." -ForegroundColor Yellow
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -NoRestart
Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -NoRestart
wsl --install

Write-Host "WSLのインストールが完了しました。システムを再起動する必要がある場合があります。" -ForegroundColor Green

# 2. アプリケーションのインストール
Write-Host "wingetで指定されたアプリケーションをインストールします..." -ForegroundColor Yellow

$appsToInstall = @(
    "Docker.DockerDesktop",
    "Git.Git",
    "JetBrains.Toolbox",
    "Microsoft.VisualStudioCode",
    "Microsoft.PowerShell", # 最新のPowerShell Core
    "Microsoft.WindowsTerminal",
    "Miro.Miro",
    "Discord.Discord",
    "Amazon.Kindle",
    "AgileBits.1Password",
    "AgileBits.1Password.CLI"
)

foreach ($app in $appsToInstall) {
    Write-Host "Installing $app..."
    winget install --id $app --accept-source-agreements --accept-package-agreements -e
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Failed to install $app. Winget exit code: $LASTEXITCODE"
    }
    else {
        Write-Host "$app installed successfully." -ForegroundColor Green
    }
}
Write-Host "アプリケーションのインストールが完了しました。" -ForegroundColor Green

# 3. GitHub ReleaseからHackGenをダウンロードして、フォントをインストール
Write-Host "HackGenフォントをダウンロードしてインストールします..." -ForegroundColor Yellow

$hackGenLatestReleaseUrl = "https://api.github.com/repos/yuru7/HackGen/releases/latest"
try {
    $releaseInfo = Invoke-RestMethod -Uri $hackGenLatestReleaseUrl
    $downloadUrl = $releaseInfo.assets | Where-Object { $_.name -like "HackGen_NF_*.zip" } | Select-Object -ExpandProperty browser_download_url -First 1

    if ($null -eq $downloadUrl) {
        throw "HackGen NFのダウンロードURLが見つかりませんでした。"
    }

    $downloadPath = "$env:TEMP\HackGen.zip"
    $extractPath = "$env:TEMP\HackGen_extracted"
    $fontsFolder = "$env:LOCALAPPDATA\Microsoft\Windows\Fonts"

    Write-Host "HackGenをダウンロード中: $downloadUrl"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -UseBasicParsing

    Write-Host "HackGenを解凍中: $downloadPath to $extractPath"
    Expand-Archive -Path $downloadPath -DestinationPath $extractPath -Force

    Write-Host "HackGenフォントをインストール中..."
    Get-ChildItem -Path $extractPath -Filter "*.ttf" | ForEach-Object {
        try {
            # Font Install Utility (Windows API) を使用してインストール
            $shell = New-Object -ComObject Shell.Application
            $shell.Namespace($fontsFolder).CopyHere($_.FullName)
            Write-Host "Installed font: $($_.Name)" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to install font $($_.Name): $($_.Exception.Message)"
        }
    }
    Remove-Item $downloadPath -Force
    Remove-Item $extractPath -Recurse -Force
    Write-Host "HackGenフォントのインストールが完了しました。" -ForegroundColor Green
}
catch {
    Write-Warning "HackGenフォントのインストール中にエラーが発生しました: $($_.Exception.Message)"
}

function Wait-ManualDownloadAndInstall {
    param (
        [string]$SoftwareName,
        [string]$DownloadUrl
    )
    Write-Host "${SoftwareName}のダウンロードページ: $DownloadUrl" -ForegroundColor Cyan
    Write-Host "上記のURLをブラウザで開き、手動で最新版をダウンロード・インストールしてください。" -ForegroundColor Cyan
    do {
        $userInput = Read-Host "${SoftwareName}のインストールが完了したら 'y' を入力してください"
    } while ($userInput -ne 'y')
}

# 4. Realforce Connectをインストール
Write-Host "Realforce Connectをダウンロードします..." -ForegroundColor Yellow
Wait-ManualDownloadAndInstall -SoftwareName "Realforce Connect" -DownloadUrl "https://www.realforce.co.jp/support/download/software/"

# 5. WL-UG69DK1のドライバをインストール
Write-Host "WL-UG69DK1のドライバをダウンロードします..." -ForegroundColor Yellow
Wait-ManualDownloadAndInstall -SoftwareName "WL-UG69DK1のドライバ" -DownloadUrl "https://www.synaptics.com/products/displaylink-graphics/downloads/windows"

# 6. Stream Deckのドライバをインストール
Write-Host "Stream Deckのドライバをダウンロードします..." -ForegroundColor Yellow
Wait-ManualDownloadAndInstall -SoftwareName "Stream Deckのドライバ" -DownloadUrl "https://elgato.com/download"

function Wait-Action {
    param (
        [string]$action
    )
    Write-Host $action -ForegroundColor Cyan
    do {
        $userInput = Read-Host "完了したら 'y' を入力してください"
    } while ($userInput -ne 'y')
}

# 7. 1Passwordにログインする
Write-Host "1Password CLIのセットアップを開始します..." -ForegroundColor Yellow
Wait-Action -action "1Passwordにログインしてください。"
Wait-Action -action "1PasswordでSSHエージェントを開始してください。"
Wait-Action -action "1Passwordで1PasswordCLIとの連携を有効にしてください。"
Wait-Action -action "1PasswordでGitコミットへの署名を有効にしてください。"

Write-Host "===== Windows dotsfiles 設定スクリプトが完了しました =====" -ForegroundColor Green
