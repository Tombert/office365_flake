# Rebuilding `winewayland.so`

`winewayland.so` is GE-Proton11-7's Wayland driver with six changes:

* `client-surface-zorder.patch`: stacks the subsurfaces of self-presenting child windows in Win32
  z-order (OneNote's lists, Excel's grid), and again whenever a child window's z-order changes
  (Excel's in-cell editor is raised above the grid only after it is shown).
* `virtual-modifiers.patch`: Shift, Ctrl and Alt that exist only in the compositor's modifier state.
* `touch-as-mouse.patch`: touchscreen contacts drive the left mouse button instead of `WM_POINTER`
  messages (one tap froze PowerPoint).
* `maximized-frame.patch`: a maximized window extends past the compositor's area by its frame, as
  on Windows, and the surface geometry leaves the frame out (a dead strip along the top and left of
  maximized Office windows on KDE Plasma).
* `maximized-resize.patch` (needs `maximized-frame.patch`): a window maximized before the
  compositor's first configure is resized to the configured size before that configure is
  acknowledged (Word's status bar was behind KDE Plasma's panel).
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

## Touch

GE's driver turns touchscreen contacts into `WM_POINTERDOWN`/`UPDATE`/`UP` hardware messages, and
Proton's wineserver queues a `WM_POINTERENTER` before the down and a `WM_POINTERLEAVE` after the up.
Windows sends those two to the window procedure; it never queues them. PowerPoint peeks the queued
`WM_POINTERENTER` without removing it, over and over (32,000 times in 1.5 s in a `+msg` trace), so
no mouse or keyboard input reaches it again and its last cursor stays on screen until it is closed.
Office would not get far with the pointer messages anyway: `GetPointerInfo` is a stub,
`GetPointerFrameInfoHistory` is not exported, the interaction context is a stub, and
`GetPointerDevices` reports no devices. The driver also sent the Wayland touch id as the pointer id
(0 for the first finger, which Windows never uses) and dropped touch motion that actually moved.

`touch-as-mouse.patch` makes the first contact drive the mouse instead: a move to the contact and the
left button down, moves while it slides, the button up when it lifts (also on `wl_touch.cancel` and
when the seat loses touch). Contacts that start while it is down are ignored. The touch-down serial
is kept as the popup grab serial, so a tap opens ribbon dropdowns. Taps, double taps (word
selection), touch drags (text selection) and dropdowns work in PowerPoint; multi-touch gestures
such as pinch zoom do not exist this way.

## Maximized windows

Windows maximizes a window with a resize frame to the work area grown by that frame on every side
(`get_min_max_info`: position `-frame`, size `work area + 2 * frame`), so the frame is off-screen.
Office draws its own title bar and frame and gives a maximized window a window region without the
part outside its monitor, which is that frame (5 px at 144 dpi). GE's driver instead made the whole
Win32 window exactly the configured maximized size (`WAYLAND_GetWindowMaxTrackSize` caps it there,
and `wayland_configure_window` sizes it there) and set the window geometry at (0, 0), so the clipped
frame was inside the geometry: a transparent strip, 3 px at 1.5x scale, along the top and left of
maximized Word on KDE Plasma, where clicks reached the window behind it. Flinging the pointer to the
top edge or a corner, to the title bar or the close button, lands there.

`maximized-frame.patch` computes that frame the way win32u does (`AdjustWindowRect` of the style
without `WS_BORDER`, at the system dpi), for toplevels with `WS_THICKFRAME` whose frame the
compositor does not hide (no server-side decorations). Maximized, the max track size and the size
`wayland_configure_window` sets include it, and the window geometry starts at the frame, like the
existing case for a fullscreen client's thin frame. The surface beyond the geometry is the frame,
which such applications leave transparent and without input, and on Windows is off-screen too.

GE's driver also moves a toplevel's Win32 position to the origin of its output whenever it refreshes
the window state (output enter and leave, focus changes). The patch moves a maximized window to that
origin minus the frame instead; at the origin itself Office's region would lose its left and top
inset and stop 5 px short of the right edge, which moves the dead strip there.

A maximized configure is acknowledged straight from the compositor's request, without the Win32
resize, when the window is at least that large. Office starts maximized before its first configure,
which Win32 sizes for the whole output (Wine's work area under Wayland), so Word's status bar stayed
below the geometry, behind the panel (also with GE's own driver). `maximized-resize.patch` makes that
shortcut need exactly the configured size plus the frame for windows with the frame; otherwise
`wayland_configure_window` resizes the window first.

At 1.5x the 5 px frame is 3 1/3 logical pixels and the geometry offset is 3, so the frame can leave
half a physical pixel at the edges; input reaches the window there.

## Cursor shapes

GE's `wayland_pointer.c` maps `OCR_SIZE` and `OCR_SIZEALL` (the four-way arrow Office shows over
things that can be moved) to `WP_CURSOR_SHAPE_DEVICE_V1_SHAPE_ALL_RESIZE`, which is new in version 2
of cursor-shape-v1; the driver sends it only when the compositor offers version 2. KWin 6.4 (SteamOS
desktop mode) offers version 2 but still validates shapes against version 1's last value
(`shape > shape_zoom_out`, fixed in later KWin), so it answers `all_resize` with a protocol error
("unknown cursor shape") that ends the Wayland connection and Excel with it. `build.sh` maps both
cursors to `..._SHAPE_MOVE` (version 1, the same four arrows). Before the rebuild the shipped binary
had the same change made in place (two bytes of the cursor table, 36 to 13).
