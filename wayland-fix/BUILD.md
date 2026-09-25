# Rebuilding `winewayland.so`

`winewayland.so` is GE-Proton11-7's Wayland driver with three changes:

* `client-surface-zorder.patch`: stacks the subsurfaces of self-presenting child windows in Win32
  z-order (OneNote's lists, Excel's grid), and again whenever a child window's z-order changes
  (Excel's in-cell editor is raised above the grid only after it is shown).
* `virtual-modifiers.patch`: Shift, Ctrl and Alt that exist only in the compositor's modifier state.
* the cursor-shape edit in `build.sh`: `all_resize` sent as `move`.

`build.sh` rebuilds it: it clones GE-Proton11-7 (Wine, wine-staging and libxkbcommon only), applies
the Wine part of GE's `patches/protonprep-valve-staging.sh`, then these changes, and compiles just the
driver in the Steam Runtime SDK image GE's Makefile names (`STEAMRT_IMAGE`), with GE's compiler flags
and configure options. It needs git, patch and Docker, takes about two minutes once the SDK image is
pulled, and writes `winewayland.so` here.

Checked on 2026-09-24: built without `virtual-modifiers.patch`, the driver disassembles to the same
instructions as the one built with GE's full build (only addresses and embedded source paths differ),
so the script reproduces GE's source, toolchain and configuration.

## Virtual modifiers

Keyboards that exist only in software type some characters by putting a modifier into the
compositor's modifier state (`wl_keyboard.modifiers`) without a key event for it. Steam's on-screen
keyboard does it for `(`, `)` and the other shifted symbols: Shift in the modifier state, then the
`9` or `0` key (it sends a real Shift key for capital letters). Wayland clients that translate keys
with xkbcommon get `(`; Windows applications go by the key state, which only key events change, and
Wine's driver synchronises only Caps, Num and Scroll Lock from the modifier state (GE's
`update_mod_state`), so Excel typed `9`. The patch presses the left Shift, Ctrl or Alt key before a
key press when the compositor reports that modifier and no key for it is down, and releases it when
the compositor clears the modifier; a real key event for that modifier takes over. Worth sending to
GE and upstream Wine.

## Cursor shapes

GE's `wayland_pointer.c` maps `OCR_SIZE` and `OCR_SIZEALL` (the four-way arrow Office shows over
things that can be moved) to `WP_CURSOR_SHAPE_DEVICE_V1_SHAPE_ALL_RESIZE`, which is new in version 2
of cursor-shape-v1; the driver sends it only when the compositor offers version 2. KWin 6.4 (SteamOS
desktop mode) offers version 2 but still validates shapes against version 1's last value
(`shape > shape_zoom_out`, fixed in later KWin), so it answers `all_resize` with a protocol error
("unknown cursor shape") that ends the Wayland connection and Excel with it. `build.sh` maps both
cursors to `..._SHAPE_MOVE` (version 1, the same four arrows). Before the rebuild the shipped binary
had the same change made in place (two bytes of the cursor table, 36 to 13).
