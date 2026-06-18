$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$ProjectDir = "$PSScriptRoot\android"
$ToolsDir   = "$env:LOCALAPPDATA\PiratBuildTools"
$SdkDir     = "$ToolsDir\android-sdk"
$GradleVer  = "8.6"

function Log  { param($m) Write-Host "  $m"       -ForegroundColor Cyan  }
function Ok   { param($m) Write-Host "  [OK] $m"  -ForegroundColor Green }
function Head { param($m) Write-Host "`n== $m ==" -ForegroundColor Yellow }
function Fail { param($m) Write-Host "  [ERR] $m" -ForegroundColor Red; Read-Host "Press Enter to exit"; exit 1 }

function TryDownload {
    param($Urls, $Out, [int]$TimeoutSec = 120)
    if ($Urls -is [string]) { $Urls = @($Urls) }
    foreach ($url in $Urls) {
        Log "Trying: $url"
        try {
            Invoke-WebRequest -Uri $url -OutFile $Out -UseBasicParsing -TimeoutSec $TimeoutSec `
                -Headers @{"User-Agent"="Mozilla/5.0"}
            return $true
        } catch {
            Log "  -> failed: $($_.Exception.Message)"
        }
    }
    return $false
}

function Find-Java {
    $j = Get-Command java -ErrorAction SilentlyContinue
    if ($j) { return $j.Source }
    $patterns = @(
        "$env:ProgramFiles\Microsoft\*\bin\java.exe",
        "$env:ProgramFiles\Eclipse Adoptium\*\bin\java.exe",
        "$env:ProgramFiles\Java\*\bin\java.exe",
        "$env:ProgramFiles\Android\Android Studio\jbr\bin\java.exe"
    )
    foreach ($p in $patterns) {
        $f = Get-Item $p -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($f) { return $f.FullName }
    }
    return $null
}

# ===============================================================================
Head "1/4  Java JDK"
# ===============================================================================

$javaExe = Find-Java
if ($javaExe) {
    Ok "Java found: $javaExe"
    $env:JAVA_HOME = $javaExe | Split-Path | Split-Path
} else {
    Log "Java not found. Installing Microsoft OpenJDK 17 via winget..."
    try {
        winget install Microsoft.OpenJDK.17 --silent --accept-package-agreements --accept-source-agreements
    } catch {
        Fail "winget failed. Install JDK 17 manually: https://adoptium.net"
    }
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","User")
    $javaExe = Find-Java
    if (-not $javaExe) { Fail "JDK installed but java.exe not found. Re-run the script." }
    $env:JAVA_HOME = $javaExe | Split-Path | Split-Path
    Ok "JDK installed: $javaExe"
}
$env:Path = "$env:JAVA_HOME\bin;$env:Path"

# ===============================================================================
Head "2/4  Gradle"
# ===============================================================================
#
# Strategy: PowerShell downloads the Gradle zip (600s timeout, multiple mirrors).
# The wrapper.properties points to the local file via file:// URL so the Java
# wrapper just extracts it -- no network call from Java at all.
#

New-Item -ItemType Directory -Force -Path $ToolsDir | Out-Null

$wrapperDir   = "$ProjectDir\gradle\wrapper"
$wrapperJar   = "$wrapperDir\gradle-wrapper.jar"
$wrapperProps = "$wrapperDir\gradle-wrapper.properties"
$gradlewBat   = "$ProjectDir\gradlew.bat"
$gradleZip    = "$ToolsDir\gradle-$GradleVer-bin.zip"

New-Item -ItemType Directory -Force -Path $wrapperDir | Out-Null

# -- gradle-wrapper.jar (~60KB bootstrap jar) ----------------------------------
if (-not (Test-Path $wrapperJar)) {
    $ok = TryDownload @(
        "https://github.com/gradle/gradle/raw/v8.7.0/gradle/wrapper/gradle-wrapper.jar",
        "https://github.com/gradle/gradle/raw/v8.6.0/gradle/wrapper/gradle-wrapper.jar",
        "https://github.com/gradle/gradle/raw/master/gradle/wrapper/gradle-wrapper.jar"
    ) $wrapperJar
    if (-not $ok) { Fail "Cannot download gradle-wrapper.jar. Is GitHub accessible?" }
    Ok "gradle-wrapper.jar downloaded ($([int](Get-Item $wrapperJar).Length/1024) KB)"
} else {
    Ok "gradle-wrapper.jar already present"
}

# -- Gradle distribution zip (~130MB) - downloaded by PowerShell, not Java ----
if (-not (Test-Path $gradleZip)) {
    Log "Downloading Gradle $GradleVer (~130 MB) via PowerShell (timeout 10 min)..."
    $ok = TryDownload @(
        "https://services.gradle.org/distributions/gradle-$GradleVer-bin.zip",
        "https://repo.huaweicloud.com/gradle/gradle-$GradleVer-bin.zip",
        "https://mirrors.cloud.tencent.com/gradle/gradle-$GradleVer-bin.zip"
    ) $gradleZip 600
    if (-not $ok) { Fail "Cannot download Gradle $GradleVer from any mirror." }
    Ok "Gradle zip downloaded ($([int](Get-Item $gradleZip).Length/1024/1024) MB)"
} else {
    Ok "Gradle zip already cached ($([int](Get-Item $gradleZip).Length/1024/1024) MB)"
}

# -- gradle-wrapper.properties pointing to local file:// URL ------------------
$gradleFwdSlash = $gradleZip -replace "\\", "/"
$distUrlEscaped = ("file:///$gradleFwdSlash") -replace ":", "\:"
@"
distributionBase=GRADLE_USER_HOME
distributionPath=wrapper/dists
distributionUrl=$distUrlEscaped
networkTimeout=120000
validateDistributionUrl=false
zipStoreBase=GRADLE_USER_HOME
zipStorePath=wrapper/dists
"@ | Set-Content -Path $wrapperProps -Encoding UTF8

# -- gradlew.bat ---------------------------------------------------------------
@'
@if "%DEBUG%"=="" @echo off
setlocal
set DIRNAME=%~dp0
if "%DIRNAME%"=="" set DIRNAME=.
set APP_HOME=%DIRNAME%
for %%i in ("%APP_HOME%") do set APP_HOME=%%~fi
set DEFAULT_JVM_OPTS="-Xmx64m" "-Xms64m"
if defined JAVA_HOME goto findJavaFromJavaHome
set JAVA_EXE=java.exe
%JAVA_EXE% -version >NUL 2>&1
if %ERRORLEVEL% equ 0 goto execute
echo ERROR: JAVA_HOME is not set. 1>&2
goto fail
:findJavaFromJavaHome
set JAVA_HOME=%JAVA_HOME:"=%
set JAVA_EXE=%JAVA_HOME%/bin/java.exe
if exist "%JAVA_EXE%" goto execute
echo ERROR: JAVA_HOME is set to an invalid directory: %JAVA_HOME% 1>&2
goto fail
:execute
set CLASSPATH=%APP_HOME%\gradle\wrapper\gradle-wrapper.jar
"%JAVA_EXE%" %DEFAULT_JVM_OPTS% %JAVA_OPTS% %GRADLE_OPTS% "-Dorg.gradle.appname=%APP_BASE_NAME%" -classpath "%CLASSPATH%" org.gradle.wrapper.GradleWrapperMain %*
:end
if %ERRORLEVEL% equ 0 goto mainEnd
:fail
set EXIT_CODE=%ERRORLEVEL%
if %EXIT_CODE% equ 0 set EXIT_CODE=1
exit /b %EXIT_CODE%
:mainEnd
endlocal
:omega
'@ | Set-Content -Path $gradlewBat -Encoding ASCII

Ok "Gradle wrapper ready"

# ===============================================================================
Head "3/4  Android SDK"
# ===============================================================================

$sdkManager = "$SdkDir\cmdline-tools\latest\bin\sdkmanager.bat"

if (-not (Test-Path $sdkManager)) {
    New-Item -ItemType Directory -Force -Path "$SdkDir\cmdline-tools" | Out-Null
    $zip = "$env:TEMP\cmdtools.zip"
    $ok = TryDownload @(
        "https://dl.google.com/android/repository/commandlinetools-win-11076708_latest.zip"
    ) $zip
    if (-not $ok) { Fail "Cannot download Android SDK tools. Check internet connection." }
    Log "Extracting SDK tools..."
    $tmp = "$env:TEMP\cmdtools_extract"
    if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
    Expand-Archive $zip -DestinationPath $tmp -Force
    Move-Item "$tmp\cmdline-tools" "$SdkDir\cmdline-tools\latest" -Force
    Remove-Item $tmp -Recurse -Force
    Remove-Item $zip -Force
    Ok "SDK tools ready"
} else {
    Ok "SDK tools already installed"
}

$platformDir   = "$SdkDir\platforms\android-34"
$buildToolsDir = "$SdkDir\build-tools\34.0.0"

if (-not (Test-Path $platformDir) -or -not (Test-Path $buildToolsDir)) {
    Log "Accepting SDK licenses..."
    ("y`n" * 20) | & $sdkManager --sdk_root="$SdkDir" --licenses 2>&1 | Out-Null
    Log "Installing SDK components (may take a few minutes)..."
    & $sdkManager --sdk_root="$SdkDir" "platforms;android-34" "build-tools;34.0.0"
    if ($LASTEXITCODE -ne 0) { Fail "SDK install failed" }
    Ok "SDK components installed"
} else {
    Ok "SDK components already installed"
}

# ===============================================================================
Head "4/4  Build APK"
# ===============================================================================

$sdkPath = $SdkDir -replace "\\", "/"
# Write without BOM - Gradle's properties parser chokes on UTF-8 BOM
[System.IO.File]::WriteAllText(
    "$ProjectDir\local.properties",
    "sdk.dir=$sdkPath`n",
    [System.Text.UTF8Encoding]::new($false)
)
Ok "local.properties -> $sdkPath"

# Also set env var so Gradle finds SDK even if local.properties parsing fails
$env:ANDROID_HOME = $SdkDir
$env:ANDROID_SDK_ROOT = $SdkDir

Log "Building APK..."
Push-Location $ProjectDir
& .\gradlew.bat assembleDebug --no-daemon
$code = $LASTEXITCODE
Pop-Location

if ($code -ne 0) { Fail "Build failed with code $code. See log above." }

$apk = Get-ChildItem "$ProjectDir\app\build\outputs\apk\debug\*.apk" -ErrorAction SilentlyContinue |
       Select-Object -First 1

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
Write-Host "  |  BUILD SUCCESSFUL                                |" -ForegroundColor Green
Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
if ($apk) {
    Write-Host "  |  $($apk.Name)" -ForegroundColor Green
    Write-Host "  |  $($apk.DirectoryName)" -ForegroundColor Green
}
Write-Host "  |                                                  |" -ForegroundColor Green
Write-Host "  |  Copy APK to phone and install.                 |" -ForegroundColor Green
Write-Host "  |  Settings > Security > Unknown sources > Allow  |" -ForegroundColor Green
Write-Host "  +--------------------------------------------------+" -ForegroundColor Green
Write-Host ""

if ($apk) { explorer.exe /select,"$($apk.FullName)" }
Read-Host "Press Enter to exit"
