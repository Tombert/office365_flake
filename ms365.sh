# ms365 -- install and run Microsoft 365 (click-to-run Office) through umu + Proton.
# Sourced into a writeShellApplication wrapper (set -euo pipefail already on).
# MS365_PROTON_GE / MS365_PROTOSODA are injected by the Nix wrapper.

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
  # Office wants an X11 surface; GE-Proton 11 ships winewayland too, prefer X11 unless overridden.
  export PROTON_USE_X11_EXCLUSIVE="${PROTON_USE_X11_EXCLUSIVE:-1}"
  mkdir -p "$MS365_PREFIX" "$LOG_DIR"
  # MS365_DEBUG=1 turns on Proton's wine log (+seh etc.) -> $LOG_DIR/steam-<appid>.log
  if [ "${MS365_DEBUG:-0}" != 0 ]; then
    export PROTON_LOG=1 PROTON_LOG_DIR="$LOG_DIR"
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

apply_registry() {
  mkdir -p "$ODT_DIR"
  local reg="$ODT_DIR/ms365.reg"
  cat > "$reg" <<'REG'
Windows Registry Editor Version 5.00

; Force Direct2D 1.0 factory: Office's ribbon/text rendering trips on newer D2D/DWrite paths in Wine.
[HKEY_CURRENT_USER\Software\Wine\Direct2D]
"max_version_factory"=dword:00000000

; Office C2R checks for the SLC/SPP licensing platform; mark it present so the installer proceeds.
[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform]
"Version"="10.0.19041.1"

; Wine's WinRT PackageManager (appxdeploymentclient) is a stub that returns E_NOINTERFACE for the
; IPackageManager revision Office asks for, and OfficeClickToRun then dereferences NULL while staging
; MSIX add-ons in its LastRun task. With the library unavailable Office logs an error and moves on.
[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"appxdeploymentclient"=""

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
REG
  if [ "${MS365_SCA:-1}" != 0 ]; then
    cat >> "$reg" <<'REG'

; Shared Computer Activation: license through a signed-in account token instead of SPP product
; activation (Wine has no Software Protection service). MS365_SCA=0 leaves this off.
[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration]
"SharedComputerLicensing"="1"
REG
  fi
  log "Applying registry tweaks"
  umu regedit /S "$reg"
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
  local reg="$ODT_DIR/shim-overrides.reg"
  mkdir -p "$ODT_DIR"
  cat > "$reg" <<'REG'
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"sppc"="native"
"ole32"="native,builtin"

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
  printf 'sppc,ole32,uia' > "$MS365_PREFIX/.ms365-shims"
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

post_install_fixups() {
  mirror_vfs
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
    log "Try:  ms365 run word"
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
  [ -f "$path" ] || die "$exe not installed (expected $path). Run: ms365 install"
  # After a crash Office offers "safe mode" via a modal prompt on the next start. Crashes are a fact
  # of life under Wine, so clear that flag unless MS365_KEEP_SAFEMODE_PROMPT=1. Only spend an extra
  # umu round-trip when the key is actually present in the hive.
  local regname; regname="$(app_regname "$app")"
  local hivekey; hivekey=$(printf 'Office\\\\16.0\\\\%s\\\\Resiliency' "$regname")   # hives escape backslashes
  if [ "${MS365_KEEP_SAFEMODE_PROMPT:-0}" = 0 ] && grep -aqF "$hivekey" "$root/user.reg" 2>/dev/null; then
    log "clearing $regname safe-mode prompt from the last crash"
    umu reg delete "HKCU\\Software\\Microsoft\\Office\\16.0\\$regname\\Resiliency" /f >/dev/null 2>&1 || true
  fi
  # a Proton version bump re-links system32; make sure the shims are still in place (no umu call here,
  # the registry overrides persist, only the files need re-checking)
  if [ "$MS365_EDITION" = 64 ]; then
    put_dll "$MS365_SPPC_SHIM" sppc.dll
    put_dll "$(runner_path)/files/lib/wine/x86_64-windows/ole32.dll" ole32_wine.dll
    put_dll "$(ole32_shim_for_runner)" ole32.dll
    put_dll "$MS365_UIA_SHIM" ms365uia.dll
    # a prefix update rewrites Wine's own class registrations; put ours back when they are gone
    if ! grep -aqF 'ms365uia' "$root/system.reg" 2>/dev/null; then
      log "re-registering shim COM classes"
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
  [ "${1:-}" = "--yes" ] || die "this deletes $MS365_PREFIX ; rerun with: ms365 reset --yes"
  cmd_kill
  rm -rf "$MS365_PREFIX"
  log "prefix removed (ODT payload in $ODT_DIR kept; delete it yourself if you want a clean download)"
}

usage() {
  cat <<USG
ms365 -- Microsoft 365 through umu-launcher + Proton

  ms365 install [download|configure|fixup]
                                      create prefix, winetricks, fetch ODT, download + install Office
                                      (a single phase reruns just that step; fixup = post-install DLL copies)
  ms365 run <app> [files...]          word excel powerpoint outlook onenote access publisher
  ms365 winetricks <verbs...>         run winetricks verbs in the prefix
  ms365 exec <exe|cmd> [args...]      run any exe (or wine builtin: cmd, regedit, winecfg, control)
  ms365 status                        show configuration and what is installed
  ms365 kill                          stop the wineserver for this prefix
  ms365 reset --yes                   delete the prefix

Environment (all optional):
  MS365_RUNNER    ge (GE-Proton from nixpkgs, default) | protosoda (Bottles Soda core) | /abs/path
  MS365_PREFIX    prefix directory  (default ~/.local/share/ms365/prefix)
  MS365_PRODUCT   ODT product id    (default O365ProPlusRetail)
  MS365_CHANNEL   update channel    (default Current)
  MS365_VERSION   pin an Office build, e.g. 16.0.18129.20158
  MS365_EDITION   64 | 32           (default 64)
  MS365_LANG      language id       (default en-us)
  MS365_EXCLUDE   ExcludeApp ids    (default "Teams OneDrive Lync Bing Groove")
  MS365_WINETRICKS verbs applied before install (default "corefonts msxml6 riched20 gdiplus")
  MS365_ODT_SETUP path to a local ODT setup.exe instead of downloading
  MS365_DEBUG=1   write Proton/Wine debug log to ~/.local/share/ms365/logs/
  MS365_KEEP_SAFEMODE_PROMPT=1  don't auto-clear Office's "start in safe mode?" prompt after a crash
  MS365_APP_ARGS  override the default per-app switches (Word: /q = no splash screen); set to "" to disable
  MS365_SCA=0     don't enable Shared Computer Activation (account-token licensing) in the prefix
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
