#!/usr/bin/env bash
# Rebuilds wayland-fix/winewayland.so: GE-Proton11-7's Wayland driver with the changes in this
# directory, compiled in Valve's Steam Runtime SDK container (the toolchain and libraries GE builds
# with, so it loads next to the rest of GE's Wine). Only the driver is built, not all of GE: about
# ten minutes and a few GB (Wine source; the SDK image is 9 GB if Docker does not have it yet).
#
#   wayland-fix/build.sh [workdir]        (default: ~/.cache/ge-wayland-build)
#
# Needs git, patch, Docker. Writes wayland-fix/winewayland.so.
set -euo pipefail

here=$(cd "$(dirname "$0")" && pwd)
B=$(mkdir -p "${1:-$HOME/.cache/ge-wayland-build}" && cd "${1:-$HOME/.cache/ge-wayland-build}" && pwd)
TAG=GE-Proton11-7
IMG=registry.gitlab.steamos.cloud/proton/steamrt4/sdk/x86_64:4.0.20260714.251823-0  # STEAMRT_IMAGE in GE's Makefile.in
W=$B/proton-ge-custom/wine

run() {  # run <dir> <script>: in the SDK container, as this user, with the work dir mounted
    docker run --rm -v "$B:$B" -w "$1" -e HOME="$B" -e CCACHE_DISABLE=1 -e XDG_CACHE_HOME="$B/xdg-cache" \
        -u "$(id -u):$(id -g)" "$IMG" bash -euc "$2"
}

echo "== GE-Proton source ($TAG)"
cd "$B"
[ -d proton-ge-custom ] || git clone -q --depth 1 --branch "$TAG" https://github.com/GloriousEggroll/proton-ge-custom
cd proton-ge-custom
git submodule update --init --depth 1 wine wine-staging libxkbcommon
# GE's patch script reverts this commit, which a shallow clone does not have
git -C wine fetch -q --depth 2 origin e813ca5771658b00875924ab88d525322e50d39f

echo "== GE's Wine patches"
# The Wine part of GE's patch script (it starts with git reset --hard and git clean in wine/)
script=patches/protonprep-valve-staging.sh
start=$(grep -n '^    pushd wine$' "$script" | head -1 | cut -d: -f1)
end=$(awk -v s="$start" 'NR > s && /^    popd$/ { n = NR } END { print n }' "$script")
{ sed -n 1,15p "$script"; sed -n "${start},${end}p" "$script"; } > "$B/wine-prep.sh"
bash "$B/wine-prep.sh" > "$B/wine-prep.log" 2>&1
if grep -q 'FAILED\|can.t find file\|Reversed (or previously applied)' "$B/wine-prep.log"; then
    echo "GE's patches did not apply cleanly, see $B/wine-prep.log" >&2; exit 1
fi

echo "== our changes"
cd "$W"
patch -p1 < "$here/client-surface-zorder.patch"
patch -p1 < "$here/virtual-modifiers.patch"
# cursor shapes: all_resize -> move (see BUILD.md)
sed -i 's/\(OCR_SIZE\(ALL\)\?, *\)WP_CURSOR_SHAPE_DEVICE_V1_SHAPE_ALL_RESIZE/\1WP_CURSOR_SHAPE_DEVICE_V1_SHAPE_MOVE/' \
    dlls/winewayland.drv/wayland_pointer.c
[ "$(grep -c 'WP_CURSOR_SHAPE_DEVICE_V1_SHAPE_ALL_RESIZE}' dlls/winewayland.drv/wayland_pointer.c)" = 0 ] ||
    { echo "cursor-shape change did not apply" >&2; exit 1; }

echo "== libxkbcommon with xkbregistry (GE builds its own; the SDK has no xkbregistry)"
run "$B" "rm -rf obj-xkbcommon deps
    meson setup obj-xkbcommon proton-ge-custom/libxkbcommon --prefix=$B/deps --libdir=lib --buildtype=release \
        -Denable-xkbregistry=true -Denable-docs=false -Denable-x11=false -Denable-bash-completion=false \
        -Denable-wayland=false -Denable-tools=false >/dev/null
    ninja -C obj-xkbcommon install >/dev/null"

echo "== generated sources (as GE's make rules do)"
run "$W" "dlls/winevulkan/make_vulkan -x vk.xml -X video.xml >/dev/null
    tools/make_specfiles >/dev/null || true
    autoreconf -fi
    tools/make_requests >/dev/null"

echo "== configure and build winewayland.so"
run "$B" "rm -rf obj-wine64 && mkdir obj-wine64 && cd obj-wine64
    export CC=x86_64-linux-gnu-gcc AR=x86_64-linux-gnu-ar RANLIB=x86_64-linux-gnu-ranlib LD=x86_64-linux-gnu-ld
    export PKG_CONFIG_PATH=$B/deps/lib/pkgconfig LDFLAGS=-L$B/deps/lib
    export CFLAGS='-Wno-discarded-qualifiers -mcmodel=small -march=nocona -mtune=core-avx2 -mfpmath=sse -O2 -fwrapv -fno-strict-aliasing -ggdb -ffunction-sections -fdata-sections -fno-omit-frame-pointer'
    $W/configure -C --prefix=$B/dst-wine64 --host=x86_64-linux-gnu --enable-werror --with-wayland \
        --with-mingw=gcc --disable-tests --enable-archs=x86_64,i386 > configure.log 2>&1 ||
        { tail -20 configure.log; exit 1; }
    make -j\$(nproc) dlls/winewayland.drv/winewayland.so > build.log 2>&1 || { tail -30 build.log; exit 1; }
    strip -o $B/winewayland.so dlls/winewayland.drv/winewayland.so"

install -m755 "$B/winewayland.so" "$here/winewayland.so"
echo "== done: $here/winewayland.so ($(sha256sum "$here/winewayland.so" | cut -c1-16))"
