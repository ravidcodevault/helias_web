param(
    [Parameter(Mandatory=$true)][ValidateSet('web','backend')][string]$Kind,
    [Parameter(Mandatory=$true)][string]$OutputRoot
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
if (Test-Path -LiteralPath $OutputRoot) { throw 'Build output must be a new directory.' }
$publish = Join-Path $OutputRoot 'publish'
$package = Join-Path $OutputRoot 'package'
New-Item -ItemType Directory -Path $publish,$package | Out-Null
Push-Location -LiteralPath $repo
try {
    if ($Kind -eq 'web') {
        dotnet publish OYA_Web_Application.csproj --configuration Release --runtime win-x64 --self-contained true --output $publish --nologo -p:DebugSymbols=false -p:DebugType=None
        if ($LASTEXITCODE) { throw 'Frontend publish failed.' }
        # Runtime settings belong only on the destination server.
        Get-ChildItem -LiteralPath $publish -File -Filter 'appsettings*.json' |
            Where-Object { $_.Name -ne 'appsettings.example.json' } |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName }
        Copy-Item -LiteralPath (Join-Path $repo 'ExcelTemplates') -Destination (Join-Path $publish 'ExcelTemplates') -Recurse
        $required = 'OYA_Web_Application.exe'
    } else {
        nuget restore QB_SQL_Sync.sln -NonInteractive
        if ($LASTEXITCODE) { throw 'Backend package restore failed.' }
        msbuild QB_SQL_Sync.csproj /t:Rebuild /p:Configuration=Release /p:Platform=AnyCPU /p:PlatformTarget=x64 /p:Prefer32Bit=false /p:AppConfig=App.config.example "/p:OutDir=$publish\" /p:DebugSymbols=false /p:DebugType=None /nologo /verbosity:minimal
        if ($LASTEXITCODE) { throw 'Backend publish failed.' }
        [xml]$configuration = Get-Content -LiteralPath (Join-Path $publish 'QB_SQL_Sync.exe.config') -Raw
        if (@($configuration.configuration.connectionStrings.add | Where-Object { -not [string]::IsNullOrWhiteSpace($_.connectionString) }).Count) { throw 'Runtime SQL credentials must not enter an artifact.' }
        if (@($configuration.configuration.appSettings.add | Where-Object { -not [string]::IsNullOrWhiteSpace($_.value) }).Count) { throw 'Runtime settings must not enter an artifact.' }
        Copy-Item -LiteralPath (Join-Path $repo 'App.config.example') -Destination $publish
        $required = 'QB_SQL_Sync.exe'
    }
    if (-not (Test-Path -LiteralPath (Join-Path $publish $required))) { throw 'Published executable is missing.' }
    'PUBLISH ONLY: This folder is inactive. Configure private runtime settings and external report paths locally before manual activation. This pipeline does not change IIS, restart applications, run QuickBooks or change SQL.' |
        Set-Content -LiteralPath (Join-Path $publish 'PUBLISH-ONLY.txt') -Encoding UTF8
    $manifest = @(Get-ChildItem -LiteralPath $publish -Recurse -File -Force | ForEach-Object {
        [pscustomobject]@{
            Path = $_.FullName.Substring($publish.Length + 1).Replace('\','/')
            Bytes = $_.Length
            SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        }
    })
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $package 'manifest.json') -Encoding UTF8
    [pscustomobject]@{
        Kind=$Kind; Commit=$env:GITHUB_SHA; RunId=$env:GITHUB_RUN_ID
        RunAttempt=$env:GITHUB_RUN_ATTEMPT; BuiltAtUtc=[DateTime]::UtcNow.ToString('o')
        RequiredExecutable=$required
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $package 'release.json') -Encoding UTF8
    Compress-Archive -Path (Join-Path $publish '*') -DestinationPath (Join-Path $package 'publish.zip') -CompressionLevel Optimal
    (Get-FileHash -LiteralPath (Join-Path $package 'publish.zip') -Algorithm SHA256).Hash |
        Set-Content -LiteralPath (Join-Path $package 'publish.sha256') -Encoding ASCII
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Publish-Artifact.ps1') -Destination $package
    Write-Output "Publish package ready: $package"
} finally { Pop-Location }
