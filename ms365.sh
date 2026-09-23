# ms365 -- install and run Microsoft 365 (click-to-run Office) through umu + Proton.
# Sourced into a writeShellApplication wrapper (set -euo pipefail already on).
# MS365_PROTON_GE / MS365_PROTOSODA are injected by the Nix wrapper.

: "${MS365_CLI:=ms365}"                   # command name shown in messages (the Nix wrapper sets it)
: "${MS365_HOME:=$HOME/.local/share/ms365}"
: "${MS365_PREFIX:=$MS365_HOME/prefix}"
: "${MS365_RUNNER:=ge}"                    # ge | protosoda | /path/to/any/proton/dir
: "${MS365_CHANNEL:=Current}"              # Current | MonthlyEnterprise | SemiAnnual | ...
: "${MS365_PRODUCT:=O365ProPlusRetail}"    # O365ProPlusRetail | O365BusinessRetail | O365HomePremRetail | ProPlus2024Retail ...
: "${MS365_LANG:=en-us}"
: "${MS365_VERSION:=}"                     # optional pinned build e.g. 16.0.18129.20158 (empty = latest on channel)
: "${MS365_EDITION:=64}"                   # 64 | 32
: "${MS365_EXCLUDE:=Teams OneDrive Lync Bing Groove}"  # ODT ExcludeApp IDs (space separated)
: "${MS365_ODT_URL:=https://officecdn.microsoft.com/pr/wsus/setup.exe}"
: "${MS365_WINETRICKS:=corefonts msxml6 riched20 gdiplus}"
: "${MS365_GAMEID:=0}"

OFFICE_ROOT="$MS365_PREFIX/drive_c/Program Files/Microsoft Office/root/Office16"
C2R_DIR="$MS365_PREFIX/drive_c/Program Files/Common Files/Microsoft Shared/ClickToRun"
ODT_DIR="$MS365_HOME/odt"
LOG_DIR="$MS365_HOME/logs"

log()  { printf '\033[1;34m[ms365]\033[0m %s\n' "$*" >&2; }
warn() { printf '\033[1;33m[ms365]\033[0m %s\n' "$*" >&2; }
die()  { printf '\033[1;31m[ms365]\033[0m %s\n' "$*" >&2; exit 1; }

runner_path() {
  case "$MS365_RUNNER" in
    ge|GE|proton-ge) echo "$MS365_PROTON_GE" ;;
    protosoda|soda)  echo "$MS365_PROTOSODA" ;;
    /*)              echo "$MS365_RUNNER" ;;
    *) die "MS365_RUNNER must be 'ge', 'protosoda' or an absolute path to a Proton directory" ;;
  esac
}

# Export the environment umu-run needs. Everything else is left to Proton defaults.
umu_env() {
  export WINEPREFIX="$MS365_PREFIX"
  export PROTONPATH; PROTONPATH="$(runner_path)"
  export GAMEID="$MS365_GAMEID"
  export PROTONFIXES_DISABLE=1          # protonfixes are for games; keep the prefix untouched
  export PROTON_NO_STEAM_FFMPEG=1
  export UMU_RUNTIME_UPDATE="${UMU_RUNTIME_UPDATE:-1}"
  export WINEDLLOVERRIDES="${WINEDLLOVERRIDES:-winemenubuilder.exe=d}"  # don't spam .desktop files
  # Wine's Wayland driver by default in a Wayland session (MS365_WAYLAND=0 for X11/Xwayland). Under
  # wlroots compositors (sway, ...) Office's popup menus lose X focus the moment they open and close
  # again; the Wayland driver keeps keyboard focus inside Wine.
  if [ "${MS365_WAYLAND:-1}" = 1 ] && [ -n "${WAYLAND_DISPLAY:-}" ]; then
    export PROTON_ENABLE_WAYLAND=1
    export PROTON_USE_X11_EXCLUSIVE=0
  else
    export PROTON_USE_X11_EXCLUSIVE="${PROTON_USE_X11_EXCLUSIVE:-1}"
  fi
  # MS365_RADV_DEBUG sets RADV_DEBUG for the AMD driver (default: none).
  if [ -z "${RADV_DEBUG:-}" ] && [ -n "${MS365_RADV_DEBUG:-}" ]; then
    export RADV_DEBUG="$MS365_RADV_DEBUG"
  fi
  mkdir -p "$MS365_PREFIX" "$LOG_DIR"
  # MS365_DEBUG=1 turns on Proton's wine log (+seh etc.) -> $LOG_DIR/steam-<appid>.log
  # MS365_WINEDEBUG replaces Proton's default channel list, e.g. "+timestamp,+pid,+tid,+debugstr,+reg"
  if [ "${MS365_DEBUG:-0}" != 0 ]; then
    export PROTON_LOG=1 PROTON_LOG_DIR="$LOG_DIR"
    if [ -n "${MS365_WINEDEBUG:-}" ]; then export WINEDEBUG="$MS365_WINEDEBUG"; fi
  fi
}

umu() { umu_env; umu-run "$@"; }

ensure_prefix() {
  if [ ! -f "$MS365_PREFIX/pfx/system.reg" ] && [ ! -f "$MS365_PREFIX/system.reg" ]; then
    log "Creating Proton prefix at $MS365_PREFIX with $(runner_path)"
    # umu's createprefix verb always exits 1 (ShellExecute of an empty exe) after Proton has built the prefix.
    umu createprefix || true
    [ -f "$MS365_PREFIX/system.reg" ] || [ -f "$MS365_PREFIX/pfx/system.reg" ] || die "prefix creation failed, see output above"
  fi
}

win_prefix_root() {
  # Proton keeps the actual WINEPREFIX under $STEAM_COMPAT_DATA_PATH/pfx
  if [ -d "$MS365_PREFIX/pfx" ]; then echo "$MS365_PREFIX/pfx"; else echo "$MS365_PREFIX"; fi
}

exclude_xml() {
  local out=""
  for a in $MS365_EXCLUDE; do out+="      <ExcludeApp ID=\"$a\" />"$'\n'; done
  printf '%s' "$out"
}

write_config() {
  mkdir -p "$ODT_DIR"
  local ver=""
  [ -n "$MS365_VERSION" ] && ver=" Version=\"$MS365_VERSION\""
  cat > "$ODT_DIR/configuration.xml" <<XML
<Configuration>
  <Add OfficeClientEdition="$MS365_EDITION" Channel="$MS365_CHANNEL"$ver SourcePath="C:\\odt">
    <Product ID="$MS365_PRODUCT">
      <Language ID="$MS365_LANG" />
$(exclude_xml)    </Product>
  </Add>
  <Updates Enabled="FALSE" />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="PinIconsToTaskbar" Value="FALSE" />
  <Logging Level="Standard" Path="C:\\odt\\logs" />
</Configuration>
XML
  log "Wrote $ODT_DIR/configuration.xml"
}

fetch_odt() {
  mkdir -p "$ODT_DIR"
  if [ -n "${MS365_ODT_SETUP:-}" ]; then
    log "Using user-supplied ODT setup.exe: $MS365_ODT_SETUP"
    cp -f "$MS365_ODT_SETUP" "$ODT_DIR/setup.exe"
  elif [ ! -s "$ODT_DIR/setup.exe" ]; then
    log "Downloading Office Deployment Tool from $MS365_ODT_URL"
    curl -fL --retry 3 -o "$ODT_DIR/setup.exe.part" "$MS365_ODT_URL"
    mv "$ODT_DIR/setup.exe.part" "$ODT_DIR/setup.exe"
  fi
  # Expose the ODT dir as C:\odt inside the prefix so SourcePath/Logging paths resolve.
  local root; root="$(win_prefix_root)"
  mkdir -p "$root/drive_c"
  ln -sfn "$ODT_DIR" "$root/drive_c/odt"
}

apply_tricks() {
  [ -z "$MS365_WINETRICKS" ] && return 0
  local marker="$MS365_PREFIX/.ms365-winetricks"
  if [ -f "$marker" ] && [ "$(cat "$marker")" = "$MS365_WINETRICKS" ]; then
    log "winetricks already applied ($MS365_WINETRICKS)"
    return 0
  fi
  log "winetricks: $MS365_WINETRICKS (this can take a while)"
  # shellcheck disable=SC2086
  umu winetricks -q $MS365_WINETRICKS
  printf '%s' "$MS365_WINETRICKS" > "$marker"
}

# Wine's Wayland driver exposes outputs at their physical size, so on a HiDPI panel Office renders
# tiny at 96 dpi. MS365_DPI sets Wine's dpi; unset, it follows the focused sway output's scale on
# the Wayland driver (96 * scale) and stays 96 on Xwayland, which scales by itself. Capped at 180:
# at exactly 192 dpi Word overflows its main thread's stack while building its first window (an
# Office recursion, same on GE-Proton 11-6 and 11-7); 180 renders and behaves fine.
wine_dpi() {
  if [ -n "${MS365_DPI:-}" ]; then echo "$MS365_DPI"; return; fi
  if [ "${MS365_WAYLAND:-1}" = 1 ] && [ -n "${WAYLAND_DISPLAY:-}" ] && command -v swaymsg >/dev/null 2>&1; then
    swaymsg -t get_outputs 2>/dev/null | python3 -c '
import json, sys
outs = json.load(sys.stdin)
o = next((o for o in outs if o.get("focused")), outs[0] if outs else {})
print(min(180, int(round(96 * float(o.get("scale", 1))))))' 2>/dev/null && return
  fi
  echo 96
}
apply_dpi() {
  local dpi; dpi="$(wine_dpi)"
  [ "$(cat "$MS365_PREFIX/.ms365-dpi" 2>/dev/null)" != "$dpi" ] || return 0
  log "setting Wine dpi to $dpi"
  umu reg add 'HKCU\Control Panel\Desktop' /v LogPixels /t REG_DWORD /d "$dpi" /f >/dev/null 2>&1 || return 0
  umu reg add 'HKCU\Software\Wine\Fonts' /v LogPixels /t REG_DWORD /d "$dpi" /f >/dev/null 2>&1 || true
  printf '%s' "$dpi" > "$MS365_PREFIX/.ms365-dpi"
}

# bump when apply_registry changes so existing prefixes pick the new tweaks up on the next run
REGISTRY_REV=6
SHIMS_REV=sppc,ole32,uia,d2d1,appinit   # bump when install_shims gains a DLL or an override
apply_registry() {
  mkdir -p "$ODT_DIR"
  local reg="$ODT_DIR/ms365.reg"
  cat > "$reg" <<'REG'
Windows Registry Editor Version 5.00

; No Direct2D factory version cap (older Office-on-Wine notes suggested one; with GE-Proton 11 Office's
; request for the D2D 1.1 factory succeeds and the ribbon renders with it).
[-HKEY_CURRENT_USER\Software\Wine\Direct2D]

; Office C2R checks for the SLC/SPP licensing platform; mark it present so the installer proceeds.
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform]
"Version"="10.0.19041.1"

; Wine's WinRT PackageManager (appxdeploymentclient) is a stub that returns E_NOINTERFACE for the
; IPackageManager revision Office asks for, and OfficeClickToRun then dereferences NULL while staging
; MSIX add-ons in its LastRun task. With the library unavailable Office logs an error and moves on.
; Wine's hvsimanagementapi (Windows Sandbox host) is a stub answering E_NOTIMPL; the Click-to-Run
; client turns that into a fatal exception when it applies product changes. Absent, it is skipped.
[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"appxdeploymentclient"=""
"hvsimanagementapi"=""

; The ODT bootstrapper itself (setup.exe), for consumer SKUs such as Home2024Retail, asks the WinRT
; PackageManager whether Office is installed from the Store and aborts when the class cannot be
; activated at all. Wine's stub answers that query well enough, so give it to setup.exe only.
[HKEY_CURRENT_USER\Software\Wine\AppDefaults\setup.exe\DllOverrides]
"appxdeploymentclient"="builtin"

; Verbose Windows Installer logs (MSI*.log in %TEMP%) so integrator MSI failures are diagnosable.
[HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Installer]
"Logging"="voicewarmupx"
"Debug"=dword:00000007

; Sign-in: Office's OneAuth stack wants the Windows Web Account Manager (WinRT classes Wine lacks,
; REGDB_E_CLASSNOTREG); these documented switches make it use the browser-based flow instead.
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity]
"EnableADAL"=dword:00000001
"DisableADALatopWAMOverride"=dword:00000001
"DisableAADWAM"=dword:00000001
"DisableMSAWAM"=dword:00000001
"DisableOneAuth"=dword:00000000

; The OneAuth stack honours these only from the Group Policy hives; without them personal
; (MSA) accounts still go to WAM after home-realm discovery and fail with 0x80040154.
[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Common\Identity]
"EnableADAL"=dword:00000001
"DisableADALatopWAMOverride"=dword:00000001
"DisableAADWAM"=dword:00000001
"DisableMSAWAM"=dword:00000001

[HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\16.0\Common\Identity]
"EnableADAL"=dword:00000001
"DisableADALatopWAMOverride"=dword:00000001
"DisableAADWAM"=dword:00000001
"DisableMSAWAM"=dword:00000001

; Office feature gates (ExternalFeatureOverrides): OneAuth otherwise still hands personal (MSA)
; accounts to the Web Account Manager "broker" after home-realm discovery, and hosts the sign-in
; page in Wine's old Gecko-based browser control, which cannot run Microsoft's login pages.
; With these, OneAuth uses the Edge WebView2 runtime (install it with: ms365 winetricks webview2).
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word]
"Microsoft.Office.Identity.TestGate.DisableBrokerForOneAuth"="true"
"Microsoft.Office.Identity.FG.IsWebView2ForOneAuthEnabled"="true"

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\excel]
"Microsoft.Office.Identity.TestGate.DisableBrokerForOneAuth"="true"
"Microsoft.Office.Identity.FG.IsWebView2ForOneAuthEnabled"="true"

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\powerpoint]
"Microsoft.Office.Identity.TestGate.DisableBrokerForOneAuth"="true"
"Microsoft.Office.Identity.FG.IsWebView2ForOneAuthEnabled"="true"

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\outlook]
"Microsoft.Office.Identity.TestGate.DisableBrokerForOneAuth"="true"
"Microsoft.Office.Identity.FG.IsWebView2ForOneAuthEnabled"="true"

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\onenote]
"Microsoft.Office.Identity.TestGate.DisableBrokerForOneAuth"="true"
"Microsoft.Office.Identity.FG.IsWebView2ForOneAuthEnabled"="true"
REG
  # Licensing mode. Office's legacy path validates the Software Protection Platform state at every
  # start and Word refuses to run when that fails (it always does under Wine: 0xC004E003). The
  # token-based "vNext" mode (LicensingNext = 2, what Microsoft 365 Apps use since version 1910)
  # skips that validation and licenses the product from the signed-in account instead, which is the
  # path a personal Microsoft 365 subscription takes. Shared Computer Activation (MS365_SCA=1) is the
  # other token-based mode; only business subscriptions can use it, a personal one is refused with
  # "cannot be used to activate Office in shared computer scenarios" (0x80004005).
  local sca=0; [ "${MS365_SCA:-0}" = 1 ] && sca=1
  cat >> "$reg" <<REG

[HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Office\\ClickToRun\\Configuration]
"SharedComputerLicensing"="$sca"

; A personal subscription makes Click-to-Run try to switch the installed product to O365HomePremRetail;
; that switch cannot complete under Wine (WinRT PackageManager) and leaves Office half-configured.
; The ProPlus install licenses fine with the personal subscription, so block SKU-to-SKU switching.
[HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Office\\ClickToRun\\Updates]
"UpdatesSkuToSkuBlocked"="1"

[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Licensing\\LicensingNext]
"$MS365_PRODUCT"=dword:00000002

REG
  # Optional Office theme pin (0 colorful, 3 dark gray, 4 black, 5 white); unset = leave Office's choice.
  if [ -n "${MS365_THEME:-}" ]; then
    cat >> "$reg" <<REG

[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\Common]
"UI Theme"=dword:0000000$MS365_THEME
REG
  fi
  log "Applying registry tweaks"
  umu regedit /S "$reg"
  printf '%s' "$REGISTRY_REV" > "$MS365_PREFIX/.ms365-registry"
}

# Native DLL shims (x86_64 only; see sppc/ and ole32-shim/ in the flake):
#  - sppc.dll: Wine's Software Protection Platform client is stubs; the Office integrator aborts in
#    SLInstallLicense (click-to-run error 0-2031 / 17002). Ours accepts licences and reports none.
#  - ole32.dll: forwarder that adds CoRegisterActivationFilter (mso30win32client GetProcAddress's it
#    and dereferences NULL) and, from its DllMain, patches other Wine stubs Office calls (see
#    ole32-shim/shim.c). The real builtin is kept alongside as ole32_wine.dll.
ole32_shim_for_runner() {
  case "$MS365_RUNNER" in
    protosoda|soda) echo "$MS365_OLE32_SHIM_SODA" ;;
    *)              echo "$MS365_OLE32_SHIM_GE" ;;
  esac
}

put_dll() { # put_dll <src> <system32 name>
  local root; root="$(win_prefix_root)"
  local dst="$root/drive_c/windows/system32/$2"
  if [ ! -f "$dst" ] || [ -L "$dst" ] || ! cmp -s "$1" "$dst"; then
    log "Installing $2 into system32"
    rm -f "$dst"
    cp -f "$1" "$dst"
    chmod 644 "$dst"
  fi
}

install_shims() {
  [ "$MS365_EDITION" = 64 ] || { warn "DLL shims are x86_64 only; 32-bit Office will hit the SLInstallLicense stub"; return 0; }
  put_dll "$MS365_SPPC_SHIM" sppc.dll
  put_dll "$(runner_path)/files/lib/wine/x86_64-windows/ole32.dll" ole32_wine.dll
  put_dll "$(ole32_shim_for_runner)" ole32.dll
  put_dll "$MS365_UIA_SHIM" ms365uia.dll
  put_dll "$MS365_D2D1_DLL" d2d1.dll
  # the same shim once more under its own name, loaded into every process through AppInit_DLLs:
  # Excel never delay-loads ole32, so the ole32 forwarder alone would leave it without the shim
  put_dll "$(ole32_shim_for_runner)" ms365shim.dll
  local reg="$ODT_DIR/shim-overrides.reg"
  mkdir -p "$ODT_DIR"
  cat > "$reg" <<'REG'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"sppc"="native"
"ole32"="native,builtin"
; d2d1-fix/: Direct2D from a newer Wine. Wine 11.0's d2d1 fills geometry groups without their fill
; mode, which turns the ribbon controls' border rings into solid blocks that hide the text.
"d2d1"="native"

; Load the shim into every process (Excel never touches ole32); the shim's DllMain keeps whichever
; instance loads second passive.
[HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Windows]
"AppInit_DLLs"="ms365shim.dll"
"LoadAppInit_DLLs"=dword:00000001

; CLSID_CUIAutomationRegistrar: Wine's uiautomationcore has no class object for it and Office
; dereferences the NULL result as soon as a document gets focus. ms365uia.dll (uia-shim/) provides it;
; per-user Classes take precedence over Wine's HKLM registration.
[HKEY_CURRENT_USER\Software\Classes\CLSID\{6E29FABF-9977-42D1-8D0E-CA7E61AD87E6}]
@="ms365 UIAutomationRegistrar"

[HKEY_CURRENT_USER\Software\Classes\CLSID\{6E29FABF-9977-42D1-8D0E-CA7E61AD87E6}\InprocServer32]
@="C:\\windows\\system32\\ms365uia.dll"
"ThreadingModel"="Both"

[HKEY_LOCAL_MACHINE\Software\Classes\CLSID\{6E29FABF-9977-42D1-8D0E-CA7E61AD87E6}]
@="ms365 UIAutomationRegistrar"

[HKEY_LOCAL_MACHINE\Software\Classes\CLSID\{6E29FABF-9977-42D1-8D0E-CA7E61AD87E6}\InprocServer32]
@="C:\\windows\\system32\\ms365uia.dll"
"ThreadingModel"="Both"

REG
  umu regedit /S "$reg"
  printf '%s' "$SHIMS_REV" > "$MS365_PREFIX/.ms365-shims"
}

# Office keeps most of its DLLs under root/vfs/<KnownFolder>/... and relies on the App-V ISV layer to
# redirect file access from the real paths (C:\Program Files\Common Files\...) into that tree. Under
# Wine that redirection is not reliable, so mirror the tree into place with symlinks (only where nothing
# exists yet). MS365_MIRROR_VFS=0 disables this.
mirror_vfs() {
  [ "${MS365_MIRROR_VFS:-1}" != 0 ] || return 0
  local root; root="$(win_prefix_root)"
  local vfs="$root/drive_c/Program Files/Microsoft Office/root/vfs"
  [ -d "$vfs" ] || return 0
  local n=0
  link_tree() { # link_tree <src dir> <dst dir>
    local src="$1" dst="$2" e name
    [ -d "$src" ] || return 0
    mkdir -p "$dst"
    for e in "$src"/*; do
      [ -e "$e" ] || continue
      name="${e##*/}"
      if [ ! -e "$dst/$name" ]; then
        ln -s "$e" "$dst/$name" && n=$((n+1))
      elif [ -d "$e" ] && [ -d "$dst/$name" ] && [ ! -L "$dst/$name" ]; then
        link_tree "$e" "$dst/$name"
      fi
    done
  }
  link_tree "$vfs/ProgramFilesCommonX64" "$root/drive_c/Program Files/Common Files"
  link_tree "$vfs/ProgramFilesCommonX86" "$root/drive_c/Program Files (x86)/Common Files"
  link_tree "$vfs/ProgramFilesX64"       "$root/drive_c/Program Files"
  link_tree "$vfs/ProgramFilesX86"       "$root/drive_c/Program Files (x86)"
  link_tree "$vfs/Fonts"                 "$root/drive_c/windows/Fonts"
  [ "$n" -gt 0 ] && log "Mirrored $n VFS entries into their real locations"
  return 0
}

# Office locates optional payload (proofing tools, converters, add-ins) through MSI component
# registrations that the Click-to-Run integrator writes on Windows; under Wine that step leaves the
# App-V registry hives empty and Word reports "missing proofing tools" with every file on disk.
# Rebuild them from the per-package manifests Click-to-Run keeps in ProgramData.
register_msi_components() {
  local root; root="$(win_prefix_root)"
  local c2r="$root/drive_c/ProgramData/Microsoft/ClickToRun"
  ls -d "$c2r"/\{*\} >/dev/null 2>&1 || { log "no Click-to-Run package data yet, skipping MSI component registration"; return 0; }
  local reg="$MS365_HOME/msi-components.reg"
  log "Registering Office MSI components and qualified components (proofing tools, converters)"
  python3 "$MS365_MSI_COMPONENTS" "$root/drive_c" "$reg" || return 0
  umu regedit /S "$reg"
  printf '%s' "$MSI_COMPONENTS_REV" > "$MS365_PREFIX/.ms365-msi-components"
}
MSI_COMPONENTS_REV=4   # 4: "<lang>\Normal" category = the spelling lexicon

post_install_fixups() {
  mirror_vfs
  register_msi_components
  # 64-bit analogue of the classic ruados/eylenburg fix: the app-v subsystem DLLs must sit next
  # to the Office binaries or WINWORD etc. die on startup under Wine.
  local bits="$MS365_EDITION"
  for f in "AppvIsvSubsystems$bits.dll" "C2R$bits.dll"; do
    if [ -f "$C2R_DIR/$f" ] && [ ! -f "$OFFICE_ROOT/$f" ]; then
      log "Copying $f into Office16"
      cp -f "$C2R_DIR/$f" "$OFFICE_ROOT/$f"
    fi
  done
}

app_exe() {
  case "$1" in
    word)       echo "WINWORD.EXE" ;;
    excel)      echo "EXCEL.EXE" ;;
    powerpoint) echo "POWERPNT.EXE" ;;
    outlook)    echo "OUTLOOK.EXE" ;;
    onenote)    echo "ONENOTE.EXE" ;;
    access)     echo "MSACCESS.EXE" ;;
    publisher)  echo "MSPUB.EXE" ;;
    *) die "unknown app '$1' (word excel powerpoint outlook onenote access publisher)" ;;
  esac
}

cmd_install() {
  umu_env
  ensure_prefix
  # Recompute paths now that pfx/ exists.
  local root; root="$(win_prefix_root)"
  OFFICE_ROOT="$root/drive_c/Program Files/Microsoft Office/root/Office16"
  C2R_DIR="$root/drive_c/Program Files/Common Files/Microsoft Shared/ClickToRun"
  apply_tricks
  apply_registry
  install_shims
  write_config
  fetch_odt
  local phase="${1:-all}"
  case "$phase" in all|download|configure|fixup) ;; *) die "install phase must be download, configure or fixup" ;; esac
  # setup.exe's exit status is not a reliable success signal under Proton (a successful install has
  # returned 3 here), so run each phase without -e/pipefail and judge by what landed on disk.
  if [ "$phase" = all ] || [ "$phase" = download ]; then
    log "ODT /download  (Office payload goes to $ODT_DIR/Office, several GB)"
    umu "$ODT_DIR/setup.exe" /download "C:\\odt\\configuration.xml" 2>&1 | tee "$LOG_DIR/odt-download.log" || true
  fi
  if [ "$phase" = all ] || [ "$phase" = configure ]; then
    log "ODT /configure  (click-to-run installer; leave the window alone)"
    umu "$ODT_DIR/setup.exe" /configure "C:\\odt\\configuration.xml" 2>&1 | tee "$LOG_DIR/odt-configure.log" || true
  fi
  post_install_fixups
  if [ -f "$OFFICE_ROOT/WINWORD.EXE" ]; then
    log "Office binaries present in $OFFICE_ROOT"
    log "Try:  $MS365_CLI run word"
  else
    warn "WINWORD.EXE not found after install; check $LOG_DIR and $ODT_DIR/logs"
    exit 2
  fi
}

app_regname() {
  case "$1" in
    word) echo Word ;; excel) echo Excel ;; powerpoint) echo PowerPoint ;; outlook) echo Outlook ;;
    onenote) echo OneNote ;; access) echo Access ;; publisher) echo Publisher ;;
  esac
}

cmd_run() {
  local app="${1:-word}"; shift || true
  local exe; exe="$(app_exe "$app")"
  umu_env
  local root; root="$(win_prefix_root)"
  local path="$root/drive_c/Program Files/Microsoft Office/root/Office16/$exe"
  [ -f "$path" ] || die "$exe not installed (expected $path). Run: $MS365_CLI install"
  # Office holds ClickToRunPackageLocker at mode 000 (a Windows deny-ACL that Wine maps to no Unix
  # permission bits) while it runs, and resets it on a clean exit. After a crash or a kill -- routine
  # under Wine -- it is left at 000, and the next launch cannot open it, so Click-to-Run aborts with
  # "Something went wrong" (0x5 / STATUS_ACCESS_DENIED on the locker). Heal it before every start.
  local locker="$root/drive_c/ProgramData/Microsoft/Office/ClickToRunPackageLocker"
  if [ -f "$locker" ] && [ ! -r "$locker" ]; then
    log "restoring ClickToRunPackageLocker permissions (left locked by an unclean exit)"
    chmod u+rwx,g+rwx "$locker" 2>/dev/null || true
  fi
  # After a crash Office offers "safe mode" via a modal prompt on the next start. Crashes are a fact
  # of life under Wine, so clear that flag unless MS365_KEEP_SAFEMODE_PROMPT=1. Only spend an extra
  # umu round-trip when the key is actually present in the hive.
  local regname; regname="$(app_regname "$app")"
  local hivekey; hivekey=$(printf 'Office\\\\16.0\\\\%s\\\\Resiliency' "$regname")   # hives escape backslashes
  if [ "${MS365_KEEP_SAFEMODE_PROMPT:-0}" = 0 ] && grep -aqF "$hivekey" "$root/user.reg" 2>/dev/null; then
    log "clearing $regname safe-mode prompt from the last crash"
    umu reg delete "HKCU\\Software\\Microsoft\\Office\\16.0\\$regname\\Resiliency" /f >/dev/null 2>&1 || true
  fi
  # registry tweaks added after the prefix was installed
  if [ "$(cat "$MS365_PREFIX/.ms365-registry" 2>/dev/null)" != "$REGISTRY_REV" ]; then apply_registry; fi
  # Office occasionally resets the product's vNext licensing flag (seen after a failed SKU switch);
  # without it the legacy validation runs and Word starts unlicensed. Re-assert it when missing.
  if ! grep -aqF "\"$MS365_PRODUCT\"=dword:00000002" "$root/user.reg" 2>/dev/null; then
    log "restoring vNext licensing mode for $MS365_PRODUCT"
    apply_registry
  fi
  if [ "$(cat "$MS365_PREFIX/.ms365-msi-components" 2>/dev/null)" != "$MSI_COMPONENTS_REV" ]; then register_msi_components; fi
  apply_dpi
  # a Proton version bump re-links system32; make sure the shims are still in place (no umu call here,
  # the registry overrides persist, only the files need re-checking)
  if [ "$MS365_EDITION" = 64 ]; then
    put_dll "$MS365_SPPC_SHIM" sppc.dll
    put_dll "$(runner_path)/files/lib/wine/x86_64-windows/ole32.dll" ole32_wine.dll
    put_dll "$(ole32_shim_for_runner)" ole32.dll
    put_dll "$MS365_UIA_SHIM" ms365uia.dll
    put_dll "$MS365_D2D1_DLL" d2d1.dll
    put_dll "$(ole32_shim_for_runner)" ms365shim.dll
    # a prefix update rewrites Wine's own class registrations; put ours back when they are gone.
    # Also re-run when this version added a DLL override the prefix does not have yet.
    if ! grep -aqF 'ms365uia' "$root/system.reg" 2>/dev/null || [ "$(cat "$MS365_PREFIX/.ms365-shims" 2>/dev/null)" != "$SHIMS_REV" ]; then
      log "re-registering shim DLL overrides and COM classes"
      install_shims
    fi
  fi
  # Default switches: Word's /q skips the splash screen, whose thread trips a COM apartment
  # teardown race in Wine (null deref in combase) that takes the whole app down.
  local defargs=""
  case "$app" in word) defargs="/q" ;; esac
  # shellcheck disable=SC2086
  exec umu-run "$path" ${MS365_APP_ARGS-$defargs} "$@"
}

cmd_winetricks() { umu_env; ensure_prefix; umu winetricks "$@"; }
cmd_exec()       { umu_env; ensure_prefix; umu "$@"; }
cmd_kill() {
  # umu can't run a bare 'wineserver -k' (it wants an exe), and the wineserver socket lives inside the
  # runtime container, so find every process whose environment carries our prefix and signal it.
  local sig n
  for sig in TERM KILL; do
    n=0
    for p in /proc/[0-9]*; do
      [ -r "$p/environ" ] || continue
      if tr '\0' '\n' 2>/dev/null < "$p/environ" | grep -q "^WINEPREFIX=$MS365_PREFIX\$"; then
        kill "-$sig" "${p#/proc/}" 2>/dev/null && n=$((n+1))
      fi
    done
    [ "$n" -eq 0 ] && break
    log "sent SIG$sig to $n process(es)"
    sleep 2
  done
}

cmd_status() {
  umu_env
  local root; root="$(win_prefix_root)"
  echo "runner      : $MS365_RUNNER -> $(runner_path)"
  echo "prefix      : $MS365_PREFIX"
  echo "odt dir     : $ODT_DIR"
  echo "product     : $MS365_PRODUCT / $MS365_CHANNEL / x$MS365_EDITION / $MS365_LANG ${MS365_VERSION:+/ pinned $MS365_VERSION}"
  echo "exclude     : $MS365_EXCLUDE"
  for a in word excel powerpoint outlook onenote access publisher; do
    local e; e="$(app_exe "$a")"
    if [ -f "$root/drive_c/Program Files/Microsoft Office/root/Office16/$e" ]; then echo "installed   : $a"; fi
  done
}

cmd_reset() {
  [ "${1:-}" = "--yes" ] || die "this deletes $MS365_PREFIX ; rerun with: $MS365_CLI reset --yes"
  cmd_kill
  rm -rf "$MS365_PREFIX"
  log "prefix removed (ODT payload in $ODT_DIR kept; delete it yourself if you want a clean download)"
}

usage() {
  cat <<USG
$MS365_CLI -- Office ($MS365_PRODUCT) through umu-launcher + Proton

  $MS365_CLI install [download|configure|fixup]
                                      create prefix, winetricks, fetch ODT, download + install Office
                                      (a single phase reruns just that step; fixup = post-install DLL copies)
  $MS365_CLI run <app> [files...]     word excel powerpoint outlook onenote access publisher
  $MS365_CLI winetricks <verbs...>    run winetricks verbs in the prefix
  $MS365_CLI exec <exe|cmd> [args...] run any exe (or wine builtin: cmd, regedit, winecfg, control)
  $MS365_CLI status                   show configuration and what is installed
  $MS365_CLI kill                     stop the wineserver for this prefix
  $MS365_CLI reset --yes              delete the prefix

Environment (all optional):
  MS365_RUNNER    ge (GE-Proton from nixpkgs, default) | protosoda (Bottles Soda core) | /abs/path
  MS365_PREFIX    prefix directory  (default $MS365_HOME/prefix)
  MS365_PRODUCT   ODT product id    (default here: $MS365_PRODUCT)
  MS365_CHANNEL   update channel    (default Current)
  MS365_VERSION   pin an Office build, e.g. 16.0.18129.20158
  MS365_EDITION   64 | 32           (default 64)
  MS365_LANG      language id       (default en-us)
  MS365_EXCLUDE   ExcludeApp ids    (default "Teams OneDrive Lync Bing Groove")
  MS365_WINETRICKS verbs applied before install (default "corefonts msxml6 riched20 gdiplus")
  MS365_ODT_SETUP path to a local ODT setup.exe instead of downloading
  MS365_DEBUG=1   write Proton/Wine debug log to $MS365_HOME/logs/
  MS365_KEEP_SAFEMODE_PROMPT=1  don't auto-clear Office's "start in safe mode?" prompt after a crash
  MS365_APP_ARGS  override the default per-app switches (Word: /q = no splash screen); set to "" to disable
  MS365_RADV_DEBUG=<flags>  AMD driver debug flags to export as RADV_DEBUG (default: none)
  MS365_THEME=<n>  pin the Office theme: 0 colorful, 3 dark gray, 4 black, 5 white
  MS365_WAYLAND=0  use X11/Xwayland instead of Wine's Wayland driver (default 1 in a Wayland session)
  MS365_DPI=<n>    Wine dpi (default: 96 * the focused sway output's scale, capped at 180, on the Wayland driver; 96 on X11)
  MS365_WINEDEBUG=<channels>  with MS365_DEBUG=1: replace Proton's Wine debug channel list
  MS365_SCA=1     use Shared Computer Activation instead of vNext licensing (business subscriptions only)
  UMU_LOG=debug   verbose umu output
USG
}

main() {
  local cmd="${1:-help}"; shift || true
  case "$cmd" in
    install)    cmd_install "$@" ;;
    run)        cmd_run "$@" ;;
    winetricks) cmd_winetricks "$@" ;;
    exec)       cmd_exec "$@" ;;
    status)     cmd_status ;;
    kill)       cmd_kill ;;
    reset)      cmd_reset "$@" ;;
    help|-h|--help) usage ;;
    *) usage; exit 1 ;;
  esac
}

main "$@"
