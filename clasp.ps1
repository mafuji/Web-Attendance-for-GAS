param(
	[Parameter(Mandatory = $true)]
	[ValidateSet("push", "pull")]
	[string]$Action,

	[Parameter(Mandatory = $true)]
	[ValidateSet("dev", "prd")]
	[string]$Environment
)

# ============================================

# 対象定義

# ============================================

$targets = @(
	@{
		Path = ".\prd\main-app"
		Config = ".clasp.$Environment.json"
		CopyToClaspJson = $true
	},
	@{
		Path = ".\prd\controller-app"
		Config = ".clasp.$Environment.json"
		CopyToClaspJson = $true
	}
)

if ($Environment -eq "dev") {

	$targets += @(
	    @{
	        Path = ".\src\main-app"
	        Config = ".clasp.json"
	        CopyToClaspJson = $false
	    },
	    @{
	        Path = ".\src\controller-app"
	        Config = ".clasp.json"
	        CopyToClaspJson = $false
	    }
	)
}

# ============================================

# 確認表示

# ============================================

Write-Host ""
Write-Host "========================================"
Write-Host "Action      : $Action"
Write-Host "Environment : $Environment"
Write-Host "========================================"
Write-Host ""

foreach ($target in $targets) {

	$configPath = Join-Path $target.Path $target.Config

	if (-not (Test-Path $configPath)) {
	    throw "設定ファイルが見つかりません: $configPath"
	}

	$config = Get-Content $configPath -Raw | ConvertFrom-Json

	Write-Host "Folder   : $($target.Path)"
	Write-Host "ScriptId : $($config.scriptId)"
	Write-Host ""
}

$answer = Read-Host "実行する場合は y を入力してください"

if ($answer -ne "y") {
	Write-Host "キャンセルしました"
	exit
}

# ============================================

# 実行

# ============================================

foreach ($target in $targets) {

	$configPath = Join-Path $target.Path $target.Config

	if ($target.CopyToClaspJson) {
	    $destJson = Join-Path $target.Path ".clasp.json"
	    Copy-Item $configPath $destJson -Force
	    Write-Host "Switched : $destJson"
	}

	Push-Location $target.Path

	try {
	    Write-Host ""
	    Write-Host "[$($target.Path)] clasp $Action 実行中..."

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
Write-Host "========================================"
Write-Host "完了しました"
Write-Host "========================================"
