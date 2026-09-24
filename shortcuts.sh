# Desktop integration for the launcher, run as `nix run .#shortcuts` (the flake prepends the store
# paths below). Works on KDE Plasma, GNOME and other freedesktop desktops, SteamOS desktop mode
# included: application-menu entries, Desktop shortcuts, the apps' own icons taken from the installed
# Office, and Word/Excel/PowerPoint as the default apps for .docx/.xlsx/.pptx. Nothing in the Office
# prefix is changed.
#
# Set by the flake: SHORTCUTS_OFFICE2024 / SHORTCUTS_MS365 (launcher packages), SHORTCUTS_ICOUTILS,
# SHORTCUTS_PYTHON, SHORTCUTS_XDG_UTILS.

usage() {
  cat <<'USAGE'
nix run .#shortcuts [-- options]

Adds menu entries, Desktop shortcuts and icons for the installed Office apps, and makes Word, Excel
and PowerPoint the default apps for .docx, .xlsx and .pptx. Install Office first
(nix run .#office2024 -- install); run it again after a git pull to point everything at the new build.

  --variant office2024|ms365   which launcher (default: office2024)
  --no-desktop                 menu entries only, no Desktop shortcuts
  --no-defaults                leave the file associations alone
  --uninstall                  remove what this added (menu entries, shortcuts, icons, file
                               associations, the launcher link); Office itself stays installed
USAGE
}

variant=office2024 desktop=1 defaults=1 uninstall=0
while [ $# -gt 0 ]; do
  case "$1" in
    --variant) variant=${2:?--variant needs office2024 or ms365}; shift ;;
    --no-desktop) desktop=0 ;;
    --no-defaults) defaults=0 ;;
    --uninstall) uninstall=1 ;;
    -h|--help) usage; exit 0 ;;
    *) echo "unknown option: $1" >&2; usage >&2; exit 1 ;;
  esac
  shift
done
case "$variant" in
  office2024) pkg=$SHORTCUTS_OFFICE2024 ;;
  ms365) pkg=$SHORTCUTS_MS365 ;;
  *) echo "--variant must be office2024 or ms365" >&2; exit 1 ;;
esac

data="${XDG_DATA_HOME:-$HOME/.local/share}"
config="${XDG_CONFIG_HOME:-$HOME/.config}"
link="$data/$variant/launcher"                  # points at the launcher build; a garbage-collection root
prefix="${MS365_PREFIX:-$data/$variant/prefix}"
office="$prefix/pfx/drive_c/Program Files/Microsoft Office/root/Office16"
apps_dir="$data/applications"
icons_dir="$data/icons/hicolor"
desk=$(xdg-user-dir DESKTOP 2>/dev/null || true); [ -n "$desk" ] && [ "$desk" != "$HOME" ] || desk="$HOME/Desktop"

log() { printf '[shortcuts] %s\n' "$*"; }

exe_of() {
  case "$1" in
    word) echo WINWORD ;; excel) echo EXCEL ;; powerpoint) echo POWERPNT ;; onenote) echo ONENOTE ;;
    outlook) echo OUTLOOK ;; access) echo MSACCESS ;; publisher) echo MSPUB ;;
  esac
}

# The file types made to open in Office by default: "extension mime app", one per line.
associations="docx application/vnd.openxmlformats-officedocument.wordprocessingml.document word
xlsx application/vnd.openxmlformats-officedocument.spreadsheetml.sheet excel
pptx application/vnd.openxmlformats-officedocument.presentationml.presentation powerpoint"

# The desktop's own xdg-mime knows its quirks; nixpkgs' copy is the fallback.
xdg_mime() {
  if command -v xdg-mime >/dev/null 2>&1; then xdg-mime "$@"; else "$SHORTCUTS_XDG_UTILS/bin/xdg-mime" "$@"; fi
}

refresh_caches() {
  if command -v update-desktop-database >/dev/null 2>&1; then update-desktop-database "$apps_dir" 2>/dev/null || true; fi
  if command -v gtk-update-icon-cache >/dev/null 2>&1; then gtk-update-icon-cache -q -t "$icons_dir" 2>/dev/null || true; fi
  if command -v kbuildsycoca6 >/dev/null 2>&1; then kbuildsycoca6 >/dev/null 2>&1 || true
  elif command -v kbuildsycoca5 >/dev/null 2>&1; then kbuildsycoca5 >/dev/null 2>&1 || true; fi
}

if [ "$uninstall" = 1 ]; then
  for f in "$apps_dir/$variant"-*.desktop "$desk/$variant"-*.desktop; do
    if [ -e "$f" ]; then rm -f "$f"; log "removed $f"; fi
  done
  find "$icons_dir" -path "*/apps/$variant-*.png" -delete 2>/dev/null || true
  # Drop our entries from the file-association lists; other apps' entries stay.
  if [ -f "$config/mimeapps.list" ]; then
    sed -i -e "s/\(=\|;\)$variant-[a-z]*\.desktop;\{0,1\}/\1/g" -e '/^[^=[]*=;*$/d' "$config/mimeapps.list"
    log "removed $variant file associations from $config/mimeapps.list"
  fi
  if [ -L "$link" ]; then rm -f "$link"; log "removed $link (nix-collect-garbage frees the space)"; fi
  refresh_caches
  log "done; Office and its prefix are untouched"
  exit 0
fi

if [ ! -f "$office/WINWORD.EXE" ] && [ ! -f "$office/EXCEL.EXE" ]; then
  echo "Office is not installed in $prefix yet; install it first: nix run .#$variant -- install" >&2
  exit 1
fi

# Menu entries point at $link rather than straight into the store, so a later run can move them to a
# newer build; registering it as a GC root keeps nix-collect-garbage from deleting the build.
mkdir -p "$(dirname "$link")"
if command -v nix-store >/dev/null 2>&1 && nix-store --add-root "$link" --realise "$pkg" >/dev/null; then :
else
  ln -sfn "$pkg" "$link"
  log "warning: could not register $link as a GC root; after nix-collect-garbage, run this again"
fi

tmp=$(mktemp -d); trap 'rm -rf "$tmp"' EXIT
mkdir -p "$apps_dir"
if [ "$desktop" = 1 ]; then mkdir -p "$desk"; fi

installed=""
for src in "$link"/share/applications/"$variant"-*.desktop; do
  name=$(basename "$src" .desktop)            # office2024-word
  app=${name#"$variant"-}                     # word
  exe=$(exe_of "$app")
  if [ -z "$exe" ] || [ ! -f "$office/$exe.EXE" ]; then log "skipping $app (not installed)"; continue; fi

  # The app's own icon: the first icon group in its EXE. Office's icons hold one PNG per size, which
  # icotool cannot select by size, so the .ico directory is read directly.
  group=$("$SHORTCUTS_ICOUTILS/bin/wrestool" -l -t 14 "$office/$exe.EXE" | head -1 | sed -n 's/.*--name=\([^ ]*\).*/\1/p')
  if [ -n "$group" ] && "$SHORTCUTS_ICOUTILS/bin/wrestool" -x -t 14 -n "$group" -o "$tmp/$app.ico" "$office/$exe.EXE" 2>/dev/null; then
    "$SHORTCUTS_PYTHON/bin/python3" - "$tmp/$app.ico" "$icons_dir" "$name" <<'EOF'
import os, struct, sys
ico, icons_dir, name = sys.argv[1:]
data = open(ico, "rb").read()
for i in range(struct.unpack_from("<H", data, 4)[0]):
    w, h, _, _, _, _, size, off = struct.unpack_from("<BBBBHHII", data, 6 + 16 * i)
    w, img = w or 256, data[off:off + size]
    if w in (16, 24, 32, 48, 64, 128, 256) and w == (h or 256) and img[:8] == b"\x89PNG\r\n\x1a\n":
        path = os.path.join(icons_dir, f"{w}x{w}", "apps", f"{name}.png")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        open(path, "wb").write(img)
EOF
  else
    log "no icon found in $exe.EXE; the menu will show a generic one"
  fi

  # Absolute Exec (menus and file managers do not see ~/.nix-profile/bin), and the Wine window class
  # so the taskbar groups Office's windows under this entry.
  sed -e "s|^Exec=.*|Exec=$link/bin/$name %F|" \
      -e "/^StartupWMClass=/d" \
      -e "\$a StartupWMClass=${exe,,}.exe" \
      "$src" > "$apps_dir/$name.desktop"
  chmod 644 "$apps_dir/$name.desktop"
  log "menu entry: $apps_dir/$name.desktop"

  if [ "$desktop" = 1 ]; then
    install -m755 "$apps_dir/$name.desktop" "$desk/$name.desktop"   # executable = trusted on KDE
    if command -v gio >/dev/null 2>&1; then gio set "$desk/$name.desktop" metadata::trusted true 2>/dev/null || true; fi
    log "desktop shortcut: $desk/$name.desktop"
  fi
  installed="$installed $app"
done

if [ "$defaults" = 1 ]; then
  mkdir -p "$config"
  while read -r ext mime app; do
    case " $installed " in
      *" $app "*)
        xdg_mime default "$variant-$app.desktop" "$mime" || true
        if [ "$(xdg_mime query default "$mime" 2>/dev/null)" = "$variant-$app.desktop" ]; then
          log "default for .$ext: $variant-$app"
        else
          log "could not make $variant-$app the default for .$ext; set it once via Open With in the file manager"
        fi ;;
    esac
  done <<< "$associations"
fi

refresh_caches
log "done. Log out and back in if the menu or the file manager does not show the changes yet."
