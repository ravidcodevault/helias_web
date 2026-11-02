param(
    [Parameter(Mandatory=$true)][ValidateSet('web','backend')][string]$Kind,
    [Parameter(Mandatory=$true)][string]$PublishRoot
)
$ErrorActionPreference = 'Stop'
$expectedHost = @{ web='AZRWUSPRDQBR02'; backend='HELIASLLC' }[$Kind]
$expectedRoot = @{ web='C:\Application'; backend='D:\DEVELOPER' }[$Kind]
$prefix = @{ web='OYA_WEB_V'; backend='OYA_BACKEND_V' }[$Kind]
$firstVersion = @{ web=8; backend=1 }[$Kind]
if ($env:COMPUTERNAME -ine $expectedHost) { throw 'Wrong deployment computer.' }
if ([IO.Path]::GetFullPath($PublishRoot).TrimEnd('\') -ine $expectedRoot) { throw 'Wrong publish root.' }
$PublishRoot = $expectedRoot
$meta = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'release.json') -Raw | ConvertFrom-Json
if ($meta.Kind -ne $Kind -or $meta.Commit -ne $env:GITHUB_SHA -or $meta.RunId -ne $env:GITHUB_RUN_ID) { throw 'Release identity mismatch.' }
$zip = Join-Path $PSScriptRoot 'publish.zip'
$expectedHash = (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'publish.sha256') -Raw).Trim()
if ($expectedHash -notmatch '^[A-Fa-f0-9]{64}$' -or (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash -ine $expectedHash) { throw 'Publish archive checksum mismatch.' }
New-Item -ItemType Directory -Path $PublishRoot -Force | Out-Null
$folders = @(Get-ChildItem -LiteralPath $PublishRoot -Directory)
$version = $firstVersion
$existing = $null
foreach ($folder in $folders) {
    if ($folder.Name -match ('^'+[regex]::Escape($prefix)+'(\d+)$')) {
        $n = 0
        if ([int]::TryParse($matches[1],[ref]$n) -and $n -ge $version) { $version=$n+1 }
        $marker = Join-Path $folder.FullName '.oya-release.json'
        if (Test-Path -LiteralPath $marker) {
            $previous = Get-Content -LiteralPath $marker -Raw | ConvertFrom-Json
            if ($previous.RunId -eq $meta.RunId -and $previous.RunAttempt -eq $meta.RunAttempt -and $previous.Commit -eq $meta.Commit) { $existing=$folder.FullName }
        }
    }
}
if ($existing) {
    Write-Output "Release already published: $existing"
    if ($env:GITHUB_OUTPUT) { "release-folder=$existing" | Add-Content -LiteralPath $env:GITHUB_OUTPUT }
    exit 0
}
$stage = Join-Path $PublishRoot ('.oya-staging-'+[Guid]::NewGuid().ToString('N'))
$destination = Join-Path $PublishRoot ($prefix+$version)
foreach ($path in @($stage,$destination)) {
    if (-not [IO.Path]::GetFullPath($path).StartsWith($PublishRoot+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Publish path escaped its root.' }
}
New-Item -ItemType Directory -Path $stage | Out-Null
[IO.File]::SetAttributes($stage, ([IO.File]::GetAttributes($stage) -bor [IO.FileAttributes]::Hidden))
# The publish account can create releases but cannot modify older application folders.
$acl = Get-Acl -LiteralPath $stage
$rule = New-Object Security.AccessControl.FileSystemAccessRule(
    [Security.Principal.WindowsIdentity]::GetCurrent().User,
    [Security.AccessControl.FileSystemRights]::Modify,
    ([Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [Security.AccessControl.InheritanceFlags]::ObjectInherit),
    [Security.AccessControl.PropagationFlags]::None,
    [Security.AccessControl.AccessControlType]::Allow)
$acl.AddAccessRule($rule)
Set-Acl -LiteralPath $stage -AclObject $acl
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::OpenRead($zip)
try {
    foreach ($entry in $archive.Entries) {
        $target = [IO.Path]::GetFullPath((Join-Path $stage $entry.FullName.Replace('/','\')))
        if (-not $target.StartsWith($stage+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe archive entry.' }
    }
} finally { $archive.Dispose() }
Expand-Archive -LiteralPath $zip -DestinationPath $stage
$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'manifest.json') -Raw | ConvertFrom-Json
$manifest = @($manifest)
$files = @(Get-ChildItem -LiteralPath $stage -Recurse -File -Force)
if ($files.Count -ne $manifest.Count) { throw 'Published file count mismatch.' }
foreach ($file in $manifest) {
    $path = [IO.Path]::GetFullPath((Join-Path $stage $file.Path.Replace('/','\')))
    if (-not $path.StartsWith($stage+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe manifest path.' }
    $actual = Get-Item -LiteralPath $path
    if ($actual.Length -ne $file.Bytes -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ine $file.SHA256) { throw 'Published file verification failed.' }
}
if (-not (Test-Path -LiteralPath (Join-Path $stage $meta.RequiredExecutable))) { throw 'Required executable missing.' }
if ($Kind -eq 'backend') {
    # Private settings are read only on HeliasLLC, after the public artifact is verified.
    $settingsFile = Join-Path $env:ProgramData 'OyaGitHubRunners\helias_backend\server-settings.json'
    try {
        $settings = Get-Content -LiteralPath $settingsFile -Raw | ConvertFrom-Json
        if ([string]::IsNullOrWhiteSpace($settings.RuntimeConfigSource) -or [string]::IsNullOrWhiteSpace($settings.QuickBooksCompanyFile)) { throw 'Missing server settings.' }
        [xml]$sourceConfig = Get-Content -LiteralPath $settings.RuntimeConfigSource -Raw
        $targetConfigFile = Join-Path $stage 'QB_SQL_Sync.exe.config'
        [xml]$targetConfig = Get-Content -LiteralPath $targetConfigFile -Raw
        foreach ($sectionName in @('connectionStrings','appSettings')) {
            $sourceSection = $sourceConfig.configuration.SelectSingleNode($sectionName)
            if ($sourceSection) {
                $copy = $targetConfig.ImportNode($sourceSection,$true)
                $old = $targetConfig.configuration.SelectSingleNode($sectionName)
                if ($old) { [void]$targetConfig.configuration.ReplaceChild($copy,$old) }
                else { [void]$targetConfig.configuration.AppendChild($copy) }
            }
        }
        $sql = $targetConfig.configuration.SelectSingleNode("connectionStrings/add[@name='MyConnectionString']")
        if (-not $sql -or [string]::IsNullOrWhiteSpace($sql.connectionString)) { throw 'Missing runtime SQL configuration.' }
        $company = $targetConfig.configuration.SelectSingleNode("appSettings/add[@key='QuickBooks.CompanyFile']")
        if (-not $company) {
            $company = $targetConfig.CreateElement('add')
            $company.SetAttribute('key','QuickBooks.CompanyFile')
            [void]$targetConfig.configuration.appSettings.AppendChild($company)
        }
        $company.SetAttribute('value',[string]$settings.QuickBooksCompanyFile)
        $targetConfig.Save($targetConfigFile)
        $meta | Add-Member -NotePropertyName RuntimeConfigSHA256 -NotePropertyValue ((Get-FileHash -LiteralPath $targetConfigFile -Algorithm SHA256).Hash)
        $meta | Add-Member -NotePropertyName RuntimeSettingsConfigured -NotePropertyValue $true
    } catch { throw 'Backend runtime settings could not be prepared. Check the private runner server-settings.json and source configuration locally; no secrets are logged.' }
}
$meta | Add-Member -NotePropertyName Version -NotePropertyValue $version
$meta | Add-Member -NotePropertyName PublishedAtUtc -NotePropertyValue ([DateTime]::UtcNow.ToString('o'))
$meta | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $stage '.oya-release.json') -Encoding UTF8
if (Test-Path -LiteralPath $destination) { throw 'Release folder already exists.' }
# Both absolute move paths were checked against the fixed publish root above.
Move-Item -LiteralPath $stage -Destination $destination
[IO.File]::SetAttributes($destination, ([IO.File]::GetAttributes($destination) -band (-bnot [IO.FileAttributes]::Hidden)))
if ($env:GITHUB_OUTPUT) { "release-folder=$destination" | Add-Content -LiteralPath $env:GITHUB_OUTPUT }
if ($env:GITHUB_STEP_SUMMARY) {
    "Published and verified: "+$destination+" (commit "+$meta.Commit+"). Live applications were not changed." | Add-Content -LiteralPath $env:GITHUB_STEP_SUMMARY
}
Write-Output "Published and verified: $destination"
