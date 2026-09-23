# Rebuilding `winewayland.so`

`winewayland.so` is GE-Proton11-7's Wayland driver built with `client-surface-zorder.patch`. The
build tree it came from was deleted to free space; this is what it took to build it on NixOS.

1. `git clone --recursive --branch GE-Proton11-7 https://github.com/GloriousEggroll/proton-ge-custom`
2. NixOS quirks (GE's build assumes `/bin/bash` and autotools on `PATH`):
   * `Makefile.in`: `SHELL := $(shell command -v bash 2>/dev/null || echo /bin/bash)`
   * `configure.sh`: `#!/usr/bin/env bash`
   * a directory of symlinks named `aclocal`, `aclocal-1.16`, `aclocal-1.17`, `autoconf`,
     `autom4te`, `automake`, `automake-1.16` pointing at nixpkgs' `autoconf`/`automake`
     binaries, put on `PATH` (GE's scripts call the versioned names).
3. `mkdir build && cd build && ../configure.sh --build-name=ge-proton-spp` (Docker as the container
   engine; the Makefile runs the whole build inside Valve's Steam Runtime SDK image), then `make`.
   GE copies its Wine submodule to `build/src-wine` and applies its own patches there.
4. Apply `client-surface-zorder.patch` to `build/src-wine/dlls/winewayland.drv/wayland_surface.c`
   and run `make` again. The driver ends up in `build/dist/files/lib/wine/x86_64-unix/winewayland.so`;
   `strip` it and copy it here.

A full `make` takes hours. Only step 4's Wine rebuild is needed after the first one.
