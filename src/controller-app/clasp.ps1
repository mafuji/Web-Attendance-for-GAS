param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("push", "pull")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [ValidateSet("dev", "prd")]
    [string]$Environment
)

$targets = @(
    ".\prd\main-app",
    ".\prd\controller-app"
)

Write-Host ""
Write-Host "========================================"
Write-Host "Action      : $Action"
Write-Host "Environment : $Environment"
Write-Host "========================================"
Write-Host ""

foreach ($target in $targets) {

    $sourceJson = Join-Path $target ".clasp.$Environment.json"

    if (-not (Test-Path $sourceJson)) {
        throw "設定ファイルが見つかりません: $sourceJson"
    }

    $config = Get-Content $sourceJson -Raw | ConvertFrom-Json

    Write-Host "Folder   : $target"
    Write-Host "ScriptId : $($config.scriptId)"
    Write-Host ""
}

$answer = Read-Host "実行する場合は y を入力してください"

if ($answer -ne "y") {
    Write-Host "キャンセルしました"
    exit
}

foreach ($target in $targets) {

    $sourceJson = Join-Path $target ".clasp.$Environment.json"
    $destJson = Join-Path $target ".clasp.json"

    Copy-Item $sourceJson $destJson -Force

    Push-Location $target

    try {

        Write-Host ""
        Write-Host "[$target] $Action 実行中..."

        if ($Action -eq "push") {
            clasp push
        }
        else {
            clasp pull
        }

        if ($LASTEXITCODE -ne 0) {
            throw "clasp $Action failed"
        }
    }
    finally {
        Pop-Location
    }
}

Write-Host ""
Write-Host "完了しました"