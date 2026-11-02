param([ValidateRange(1, 65535)][int]$Port = 5000)
$ErrorActionPreference = 'Stop'
try {
    Set-Location -LiteralPath $PSScriptRoot
    Add-Type -AssemblyName System.Security
    $secretPath = Join-Path $env:LOCALAPPDATA 'OyaWebApplication\pubnub-secret.dpapi'
    $needsSecret = [string]::IsNullOrWhiteSpace($env:PubNub__SecretKey)
    if ($needsSecret) {
        if (Test-Path -LiteralPath $secretPath) {
            try {
                $bytes = [Security.Cryptography.ProtectedData]::Unprotect(
                    [IO.File]::ReadAllBytes($secretPath), $null,
                    [Security.Cryptography.DataProtectionScope]::CurrentUser)
                try { $env:PubNub__SecretKey = [Text.Encoding]::UTF8.GetString($bytes) }
                finally { [Array]::Clear($bytes, 0, $bytes.Length) }
                Write-Host 'Loaded the saved Windows-protected PubNub credential.'
            } catch {
                throw 'Saved PubNub credential cannot be read by this account. Use the original Windows account or ask your administrator to reconfigure it.'
            }
        } else {
            Write-Host 'First run: enter the Secret Key matching this web app keyset.'
            Write-Host 'It will be encrypted for this Windows account on this server.'
            $secure = Read-Host 'PubNub Secret Key (hidden)' -AsSecureString
            if ($secure.Length -eq 0) { throw 'The Secret Key is required.' }
            $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
            try { $bytes = [Text.Encoding]::UTF8.GetBytes([Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)) }
            finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer); $secure.Dispose() }
            try {
                $directory = Split-Path $secretPath
                New-Item -ItemType Directory -Force -Path $directory | Out-Null
                $encrypted = [Security.Cryptography.ProtectedData]::Protect(
                    $bytes, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
                [IO.File]::WriteAllBytes($secretPath, $encrypted)
                $acl = New-Object Security.AccessControl.FileSecurity
                $acl.SetAccessRuleProtection($true, $false)
                foreach ($account in @([Security.Principal.WindowsIdentity]::GetCurrent().Name, 'NT AUTHORITY\SYSTEM')) {
                    $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($account, 'FullControl', 'Allow')))
                }
                Set-Acl -LiteralPath $secretPath -AclObject $acl
                $env:PubNub__SecretKey = [Text.Encoding]::UTF8.GetString($bytes)
            } finally { [Array]::Clear($bytes, 0, $bytes.Length) }
            Write-Host 'PubNub credential saved with Windows protection.' -ForegroundColor Green
        }
    }
    Write-Host "Starting http://localhost:$Port. Keep this window open."
    & (Join-Path $PSScriptRoot 'OYA_Web_Application.exe') --urls "http://localhost:$Port"
    exit $LASTEXITCODE
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}