# Microsoft 365 on Linux via umu + GE-Proton (Nix flake)

Microsoft 365 (click-to-run Office) famously does not work on stock Wine. This flake is a
best-effort attempt to run it through **umu-launcher** with **GE-Proton** (the Wine build that
Valve/GloriousEggroll ship for games) instead, with the Bottles project's **ProtoSoda** Wine core
available as a second runner. It packages nothing from Microsoft: the Office Deployment Tool and
Office itself are downloaded at install time by the `ms365` script. You need a Microsoft 365
licence to sign in.

## Where the recipe comes from

* A February 2026 report in the Bottles tracker of Office 365 x64 running on Wine 10.20 with
  `corefonts msxml6 riched20 gdiplus`, the Office Deployment Tool and a couple of DLL copies.
* The classic ruados / eylenburg Office-on-Wine notes (Direct2D registry tweak, copying the
  `AppvIsvSubsystems*` and `C2R*` DLLs next to the Office binaries).
* The September 2026 Bottles announcement that Microsoft 365 installs, signs in (with 2FA) and runs
  Word on their Soda 11 Wine core in a Windows 10 prefix. GE-Proton 11 is the same Wine 11
  bleeding-edge lineage, and ProtoSoda is that Soda core in Proton layout, so both are offered.

## Usage

```sh
nix run .#ms365 -- install        # create prefix, winetricks, fetch ODT, download + install Office
nix run .#word                    # or: nix run .#ms365 -- run word ~/doc.docx
nix run .#ms365 -- status
nix run .#ms365 -- help
```

Or install it: `nix profile install .#ms365` gives you `ms365`, `ms365-word`, `ms365-excel`, ...
plus `.desktop` entries.

The install downloads several GB into `~/.local/share/ms365/odt/Office` and then runs the
click-to-run installer inside the prefix. Leave the installer window alone; it looks frozen for
long stretches.

## Optional: Office Home 2024 (retail key)

Microsoft 365 stays the default (`nix run .`, `.#word`, `packages.default`). The flake also packages
the one-time-purchase **Office Home 2024** as a separate, opt-in output, `.#office2024`: the same
launcher with the `Home2024Retail` product (Current channel), its own commands and its own prefix in
`~/.local/share/office2024`, so the two can be installed side by side. Home 2024 is Word, Excel,
PowerPoint and OneNote.

```sh
nix run .#office2024 -- winetricks webview2   # the Microsoft sign-in page needs it
nix run .#office2024 -- install
nix run .#word2024                            # also .#excel2024 .#powerpoint2024 .#onenote2024
```

`nix profile install .#office2024` gives `office2024`, `office2024-word`, ... and "Microsoft Word 2024
(Proton)" style desktop entries. Every `MS365_*` knob below applies too.

Activation: redeem the key at <https://setup.office.com> first, which attaches the licence to your
Microsoft account. Do not type the key into Office ("I have a product key"): entering a key goes
through the Windows Software Protection Platform, which Wine does not have (the `sppc/` shim refuses
key installs). The product runs in the same vNext mode as Microsoft 365; at the first start Word shows
"Sign in to get started", and signing in with the account the key was redeemed to activates it.
Verified with a retail Home 2024 key.

**Office LTSC 2024 keys do not work.** LTSC (`ProPlus2024Volume`, `Standard2024Volume`, and the
cheap "LTSC" keys sold online) is volume licensed. It has no account-based activation: it only
activates through the Windows Software Protection Platform, with a MAK key checked against
Microsoft's activation servers or through a KMS host, and Wine does not implement that service
(`sppc/` is a minimal stand-in that stores licences and refuses key installs). Getting LTSC to
activate would mean implementing enough of SPP (`sppsvc` and the `SL*` licensing APIs Office calls)
for Office to install a key, activate it and then validate the licence. **Pull requests for that are
welcome and encouraged** if you want to give it a try.

**OneNote works** as of 2026-09-23: it starts, signs in, shows the notebook and page lists and the
page, and edits save. Getting there took four fixes (all in the table below): a start-up crash in
Office's App-V layer, page text drawn as black boxes, the navigation panes hidden behind a black
layer (a Wayland driver fix, `wayland-fix/`), and OneNote's own "How do you want to start OneNote?"
prompt after every unclean exit. OneNote for the web (installable as a PWA) remains an alternative
on the same OneDrive notebooks.

Pull requests are welcome, for anything in the known-problems list. Most fixes here
follow the same loop: `MS365_DEBUG=1` (plus `MS365_WINEDEBUG` channels), find the Wine stub or
missing class Office trips on in `~/.local/share/<ms365|office2024>/logs/steam-0.log`, and fill the
gap in one of the shims.

## Knobs

All optional, all environment variables:

| Variable | Default | Meaning |
|---|---|---|
| `MS365_RUNNER` | `ge` | `ge` (nixpkgs `proton-ge-bin`), `protosoda` (Bottles Soda core), or an absolute Proton dir |
| `MS365_PREFIX` | `~/.local/share/ms365/prefix` | Proton compat-data dir (Wine prefix lives in `pfx/`) |
| `MS365_PRODUCT` | `O365ProPlusRetail` | ODT product id (`O365BusinessRetail`, `O365HomePremRetail`, `ProPlus2024Retail`, ...) |
| `MS365_CHANNEL` | `Current` | Update channel |
| `MS365_VERSION` | unset | Pin an Office build, e.g. `16.0.18129.20158`. Older builds are less likely to trip Wine |
| `MS365_EDITION` | `64` | `64` or `32` |
| `MS365_LANG` | `en-us` | Language |
| `MS365_EXCLUDE` | `Teams OneDrive Lync Bing Groove` | ODT `ExcludeApp` ids. Add `Outlook OneNote Access Publisher` for a minimal install |
| `MS365_WINETRICKS` | `corefonts msxml6 riched20 gdiplus` | Verbs applied before install |
| `MS365_ODT_SETUP` | unset | Use a local ODT `setup.exe` instead of downloading the current one |
| `MS365_SCA` | `0` | `1` switches to Shared Computer Activation (business subscriptions only) instead of vNext token licensing |
| `MS365_RADV_DEBUG` | unset | AMD driver flags to export as `RADV_DEBUG` (debugging aid, not needed) |
| `MS365_TRACE_MODULE` | unset | debugging: the shim logs every export lookup into this DLL and every failed lookup (`MS365_DEBUG=1` to see them) |
| `MS365_WAYLAND` | `1` | Wine's Wayland driver in a Wayland session; `0` for X11/Xwayland (popup menus close instantly there under sway and other wlroots compositors) |
| `MS365_NO_PIXEL_UNITS` | unset | `1` turns off the shim's Direct2D pixel-unit-mode fix (Excel's formula-bar icons, see the fixes table) |
| `MS365_DPI` | unset | Wine dpi. Unset: 96 × the focused sway output's scale (capped at 180, see "Known problems") on the Wayland driver, 96 on X11 |
| `UMU_LOG` | unset | `1` or `debug` for umu output |

Example, try the Soda core with a minimal install:

```sh
MS365_RUNNER=protosoda MS365_EXCLUDE="Teams OneDrive Lync Bing Groove Outlook OneNote Access Publisher" \
  nix run .#ms365 -- install
```

## Debugging

* Logs: `~/.local/share/ms365/logs/` (umu/wine stderr) and `~/.local/share/ms365/odt/logs/` (ODT).
* `ms365 exec regedit`, `ms365 exec winecfg`, `ms365 exec cmd`, `ms365 winetricks <verbs>`.
* `ms365 install download` / `ms365 install configure` rerun a single phase.
* `ms365 reset --yes` deletes the prefix but keeps the downloaded Office payload.
* If a run wedges, `ms365 kill`.

## What it took (and where it stands)

Status as of 2026-09-16 with GE-Proton11-7 and Microsoft 365 Apps build 16.0.20326.20144:

* The Office Deployment Tool downloads and installs the full suite inside the prefix.
* Word starts (with `/q`, no splash screen), draws its start screen and ribbon, and shows the
  Sign in button. On first start it also raises a "Microsoft Office cannot verify the license"
  dialog because the licensing shim below reports no licences. Sign-in / activation has not been
  verified yet.
* Excel, PowerPoint, Outlook and the rest are installed but untested.

Fixes the flake applies automatically, each one found by reading the Wine and Click-to-Run logs:

| Problem | Fix |
|---|---|
| Installer error 0-2031 (17002): integrator aborts in `sppc.dll.SLInstallLicense`, a Wine stub | `sppc/`: replacement Software Protection Platform client DLL that accepts licence installs and reports nothing installed |
| Installer crash in the LastRun task: Wine's WinRT `PackageManager` returns `E_NOINTERFACE` for a newer interface and Office dereferences NULL | `appxdeploymentclient` DLL override disabled; Office logs the failure and continues |
| Excel never touches ole32, so the ole32 shim (all of the fixes below that live in it) never loaded there | the same DLL is also installed as `ms365shim.dll` and loaded into every process through `AppInit_DLLs`; whichever instance loads second stays passive |
| Word: `CoRegisterActivationFilter` missing from ole32 (mso30win32client dereferences NULL) | `ole32-shim/`: forwarder `ole32.dll` that re-exports the builtin (kept as `ole32_wine.dll`) and implements the function |
| Word: `SetFileShortNameW`, `FindPackagesByPackageFamily`, `SetThreadpoolTimerEx` not exported by Wine's kernel32; the loader binds them to aborting stubs | the ole32 shim registers a loader notification and rewrites those import slots in every module with benign replacements |
| Word: special user APCs (`QueueUserAPC2`) dispatched wrongly by this Wine, Office aborts | the shim hides `QueueUserAPC2` from `GetProcAddress`; Office uses its pre-20H1 path |
| Word: splash-screen thread crashes in combase during COM apartment teardown, Office's crash handler kills the app | Word is launched with `/q` |
| `mso.dll` and friends live in Office's VFS tree; the App-V redirection layer is unreliable under Wine | the VFS tree is mirrored into `Program Files` with symlinks |
| Office offers safe mode after every crash via a modal prompt | the launcher clears the Resiliency key before starting an app |
| Sign-in dies with 53u4r / 12009 after the password (or on the email page) | Wine's winhttp/wininet reject unimplemented option codes with 12009; the shim accepts them (winhttp 77/140, wininet 11) |
| Licensing dialog fails with E_NOINTERFACE | shim serves `ILanguageStatics` and `IJsonObjectStatics`, which Wine's WinRT factories lack |
| Word exits at start on the legacy licensing path | product set to vNext licensing mode (LicensingNext = 2), SCA off by default |
| Ribbon font/size boxes, Comments/Editing/Share buttons and the title-bar search field are solid grey blocks; a control only shows its text while hovered, icons vanish under the hover highlight | GE-Proton 11's Wine (11.0 base) ships a `d2d1` that ignores the fill mode of geometry groups. Office draws each control's border as two nested rounded rectangles with even-odd fill (a ring); Wine fills the union, and the compositor stretches that slab over the text. `d2d1-fix/` ships `d2d1.dll` from nixpkgs' Wine 11.16 (fixed upstream in February 2026) with its builtin signature blanked so Proton loads it, registered as a native override. |
| Dialogs (e.g. "save changes?") make the whole window flicker and the document area draw at the wrong scale on the Wayland driver; Excel shows black crosshair lines across the window | Office draws dialog shadows with unowned layered `MSO_BORDEREFFECT_WINDOW_CLASS` popups; the Wayland driver makes each an independent toplevel, sway tiles them and the layout reshuffles until the dialog closes. The shim gives those windows an owner (the active window), so they become transient and float |
| Thin lines across the middle of the screen on the Wayland driver (a vertical and a horizontal one through a maximized Word) | those are the same shadow strips, floated: a Wayland client cannot position its toplevels, so sway centres each one. On the Wayland driver the shim also subclasses the strips and drops `SWP_SHOWWINDOW` in `WM_WINDOWPOSCHANGING`, so they never show; Office just has no window shadows |
| Office Home 2024 install: `setup.exe` aborts half a second in, no log (exit 3) | for consumer SKUs the ODT bootstrapper asks the WinRT `PackageManager` whether Office is installed from the Store and aborts when the class cannot be activated. Wine's `appxdeploymentclient` answers that query, so it is enabled for `setup.exe` only (`AppDefaults\setup.exe\DllOverrides`); `OfficeClickToRun.exe` keeps it disabled (see the LastRun row above) |
| Word exits with code 64 a second after start, before the "Sign in to get started" dialog can open | the dialog (react-native-win32) draws its icons with `ID2D1DeviceContext5::CreateSvgDocument`, a stub returning `E_NOTIMPL` in Wine's d2d1; Office writes through the missing document. The shim wraps `D2D1CreateFactory` (Office links it as `d2d1 #1`) and on the first call patches the device-context vtable so `CreateSvgDocument` falls back to an empty document that accepts everything and draws nothing. The dialog works, its two icons are blank |
| OneNote refuses to start: "You'll need to install the Desktop Experience before you start OneNote" | OneNote checks that the Tablet PC ink object `CLSID_InkDisp` (InkObj.dll) is registered; Wine's `inkobj` is an empty stub, so OneNote concludes it is on Windows Server without the Desktop Experience feature. `ms365uia.dll` (uia-shim/) now also serves an empty `InkDisp` (every ink method `E_NOTIMPL`), registered by the launcher |
| OneNote runs `SELECT Name FROM Win32_ServerFeature`; Wine's WMI answers with an empty result where client Windows says "invalid class" | the shim wraps `CoCreateInstance(Ex)` for `CLSID_WbemLocator` only and makes that query fail with `WBEM_E_INVALID_CLASS`. OneNote delay-loads `CoCreateInstanceEx` through ntdll's resolver, which the `GetProcAddress` hook never sees, so the shim also fills those delay-load slots up front. Not needed once `InkDisp` is registered, kept as the client-Windows answer |
| OneNote's React Native JavaScript thread loops in Office's error reporting (hundreds of thousands of handled access violations a minute) | react-native-win32 builds a `Windows.Web.Http` `HttpBaseProtocolFilter` and `HttpClient` at start; Wine has no `Windows.Web.Http`. The shim serves stand-ins (filter, cache control, client, `HttpMethod`; requests fail with `0x80072EFD` "cannot connect") and exports `DllGetActivationFactory`; the launcher registers them under `WindowsRuntime\ActivatableClassId` with `ms365shim.dll` as `DllPath`. **Experimental.** (The blank navigation pane once blamed on this was the Wayland stacking problem below) |
| OneNote crashes about every second start (page fault writing 0x8 in `AppVIsvSubsystems64`, under winhttp → crypt32 → `CloseHandle`) | Wine's crypt32 builds its default certificate chain engines on first use; when two threads verify a certificate at once, both build one and the loser frees its copy immediately. Closing that copy's registry store (the key, then its change-notification event) crashes in Office's App-V layer. The shim wraps `CertGetCertificateChain` and serialises the first chain on each default engine, so there is no loser |
| OneNote opens with "How do you want to start OneNote?" (start normally / delete the cache / delete settings) after a crash or a kill | OneNote keeps its own `ConsecutiveBootCrashes` / `ConsecutiveEarlyCrashes` counters under `OneNote\General`; the launcher resets them along with Office's safe-mode prompt (kept with `MS365_KEEP_SAFEMODE_PROMPT=1`) |
| OneNote draws every line of page text (date, time, body) as a black box; only the title shows | OneNote renders each text line into a cache cell: clip, clear, clip again, clear to white, draw the glyphs, at DPI 180 with a translated world transform. Wine's d2d1 (including 11.16) turns a `PushAxisAlignedClip` rectangle into pixels as (rect × dpi/96) × transform, so the translation is not scaled while everything drawn is; at any DPI other than 96 a translated clip lands elsewhere and the line is clipped away. The shim wraps `PushAxisAlignedClip` and hands Wine a transform whose translation is already scaled while it computes the clip |
| OneNote's notebook and page lists stay black and its page is black until the window is resized; Excel's cell grid is sometimes blank until a resize | GE's Wayland driver gives each child window with its own swapchain a subsurface below the window surface and stacks it there when it is created or moved, so the one configured last ends up on top whatever the Win32 z-order says: OneNote's full-size navigation background (the bottom sibling, created after the lists) covered them, Excel's `XLDESK` covered the grid. `wayland-fix/client-surface-zorder.patch` restacks all client subsurfaces of a window in Win32 order whenever one is stacked; `wayland-fix/winewayland.so` is GE-Proton11-7's driver built with it (Steam Runtime 4 SDK container, stripped), and the flake swaps it into the runner. Worth sending upstream to GE |
| Excel's formula-bar buttons (name-box ▾, ⋮ grip, cancel / enter / fx, the expand chevron) and the splitter next to the sheet tabs are black boxes, or icons shifted down and cut in half | Office draws those icons as glyphs from its own symbol fonts (`OFFSYM*.TTF`) with Direct2D in pixel unit mode (`D2D1_UNIT_MODE_PIXELS`) on contexts whose DPI it sets to the system DPI. Wine's d2d1 (including 11.16) stores the unit mode but still scales everything by dpi / 96, so at 180 dpi each icon is drawn 1.875 times too large and too far down, mostly outside its box; the empty rest is transparent and shows as black. The shim wraps the device-context vtable: while a context is in pixel mode Wine's DPI is held at 96 and `GetDpi` reports the DPI Office set. `MS365_NO_PIXEL_UNITS=1` turns it off |
| Segoe UI text (OneNote's canvas and messages, some dialogs) renders in Times New Roman | GE-Proton's prefix template maps Segoe UI to Times New Roman. The flake fetches Selawik (Microsoft's MIT-licensed, metric-compatible stand-in for Segoe UI), the launcher installs it into `windows\Fonts` and maps the Segoe UI family (regular, Semibold, Semilight, Light) to it |
| "Missing proofing tools" banner and no spell checking although the dictionaries are installed | Office finds its proofing tools through the Windows Installer API (component paths, feature states, "qualified components" per category and language), registrations the Click-to-Run integrator never writes under Wine. `msi-components.py` rebuilds all of them from the package manifests, and the ole32 shim answers the MSI calls Office makes with an empty product code (its "whichever package owns it" convention, imported by ordinal) from the registered products |
| Ticking the Solver add-in (any add-in opened from `Library`) crashes Excel in the App-V layer | Office's mark-of-the-web check asks `virtdisk` whether the file sits on a mounted disk image (`GetStorageDependencyInformation`). Wine's stub reports success without filling the buffer and mso reads the first entry regardless. The shim answers `ERROR_VIRTDISK_NOT_VIRTUAL_DISK`, the "no backing image" case mso handles |
| "Microsoft Excel cannot access the file …\SOLVER.XLAM"; no VBA at all (macros, the editor, recording) | Office locates VBA through an MSI qualified component (`vbe.dll_7.1` in category `{D304E920-…}`). The manifests don't say which file a qualifier stands for, and `msi-components.py` used to give every unmapped qualifier the feature's first file, `msvcr100.dll` for VBA and wrong for 155 others (`1033\solver.xlam` was `SOLVER32.DLL`). Qualifiers now resolve to the file they name; the launcher re-registers when `MSI_COMPONENTS_REV` changes |
| Workbooks and add-ins with UserForms don't load (Solver: still "cannot access the file") | VBA creates forms through the MSForms designer package class in `FM20.DLL` (`{AC9F2F90-…}`), one of the COM registrations Click-to-Run publishes only into its App-V virtual registry, which Wine's COM never sees. The launcher runs `regsvr32 /s FM20.DLL` once per prefix |
| Any VBA call through a sheet or workbook code name (`Sheet1.Range(…)`, `ThisWorkbook.Worksheets(…)`) gives "Automation error" or crashes Excel; Solver's button did both | Excel's type library is 32-bit (shared with 32-bit Office) and stores vtable sizes in 4-byte slots: `_Worksheet`, 158 methods, is stored as 632. Wine's typelib reader scales every method's `oVft` to 8-byte slots but hands `TYPEATTR.cbSizeVft` out unscaled, so the interface looks 79 methods long, and VBA sizes its document-module dispatch tables from it (`Range` is method 100). The shim wraps `ITypeInfo::GetTypeAttr` on Wine's shared type-info vtable and doubles `cbSizeVft` for interfaces from 32-bit libraries. Worth sending upstream to Wine |
| The first Solve of a session ends with "Solver encountered an error value in the Objective Cell or a Constraint cell" and nothing changes; the second works | Before loading `SOLVER32.DLL`, Solver's VBA calls `GetDllDirectory` and, since 0 means both "none set" and failure, checks `Err.LastDllError`. Windows leaves `ERROR_SUCCESS` there; Wine leaves whatever the thread last failed at, and Office's C2R layer has just left `ERROR_ENVVAR_NOT_FOUND`, so Solver returns its default error code without running. The shim wraps `GetDllDirectoryA/W` (Solver reaches it through `GetProcAddress`) and clears the error when no directory is set |

Debug aids: `MS365_DEBUG=1` writes Proton's Wine log with `+seh`; `MS365_DEBUG=1 PROTON_LOG="+module"`
lists every unresolved import ("No implementation for ..."), which is how the kernel32 gaps were
found. Office's own logs land in `drive_c/users/steamuser/AppData/Local/Temp` (`PUTER-*.log` for
Click-to-Run, `Diagnostics/<APP>/` for the apps).

## Licensing and sign-in (where it stands)

Office licenses itself through the Windows Software Protection Platform (SPP), which Wine does not
have. The `sppc/` shim is a minimal stand-in: it stores the licence files the installer hands it,
serves the SKU policies from them, and lets the out-of-box 5-day Grace licence run with a persisted
timer. Nothing is reported as activated and no product keys are installed. Office's full SPP
validation still fails under Wine (0xC004E003), and on the legacy licensing path Word then refuses
to run ("Word has run into an error ... repair now?").

The flake therefore puts the product into Microsoft's token-based licensing mode ("vNext",
`HKCU\...\Common\Licensing\LicensingNext\<product> = 2`, what Microsoft 365 Apps use since
version 1910). In that mode Office skips the SPP validation, starts as "Unlicensed Product", and
licenses itself from the signed-in account through the Office Licensing Service. Shared Computer
Activation (`MS365_SCA=1`) is the alternative token mode for business subscriptions; a personal
subscription is refused there with "cannot be used to activate Office in shared computer
scenarios" (0x80004005).

Sign-in works with a personal Microsoft account: OneAuth is kept, both Web Account Manager paths
are switched off (`Common\Identity` and the Policies hives), the OneAuth broker is disabled and its
login page is rendered by the Edge WebView2 runtime (feature gates under
`ExperimentConfigs\ExternalFeatureOverrides`). The WebView2 runtime is not in the flake recipe yet:
`ms365 winetricks webview2` (680 MB) installs it into the prefix. Two request paths needed help from
the ole32 shim: Wine's winhttp and wininet reject option codes they have not implemented with
error 12009, which Office reports as sign-in error 53u4r / code 12009, so the shim accepts those
tuning options.

What has been verified: sign-in completes, Word shows the account's OneDrive documents, and Office
fetches the account's entitlements. What has not: an actual licence, because the test account had
no Microsoft 365 subscription that includes the desktop apps. In that case Office tries to open its
in-app purchase dialog, a WebView2 window in DirectComposition mode, and Wine's DirectComposition is
a stub, so Word exits with code 64 instead. (Exit 64 is Office's crash handler; the licensing dialog's
SVG icons were a confirmed cause of it for Office 2024, fixed in the shim, see the table above, so this
may have been the same crash.) Buy or manage the subscription on the web, then sign in
again in Word.

Other WinRT gaps the ole32 shim fills for the licensing code: `Windows.Globalization.Language`
statics (`ILanguageStatics`) and `Windows.Data.Json.JsonObject` statics (`Parse`/`TryParse`, built
on Wine's `JsonValue`).

## Known problems

* **Proofing categories are mapped by file role.** The Click-to-Run manifests list which qualified
  component categories a language package publishes but not which file each one stands for;
  `msi-components.py` maps them by name (speller and "Normal" dictionary to MSSP*.LEX, grammar to
  MSGR*.LEX, hyphenation and thesaurus split between engine DLL and lexicon). Spelling works; if
  hyphenation or the thesaurus (Shift+F7) refuse a language, those two mappings are the suspects.
* **Right-click and other popup menus close immediately under sway on the X11 driver**
  (`MS365_WAYLAND=0`). Office activates its popup, hides and re-shows it while positioning it, and
  sway's Xwayland layer moves keyboard focus back to the main window in between; Wine then sends
  Office the message that cancels the menu. The Wayland driver, the default, keeps focus inside Wine.
* **Wayland driver and HiDPI.** The driver reports the monitor at 96 dpi, so the launcher sets Wine's
  dpi from the sway output scale. At exactly 192 dpi Word overflows its main thread's stack while
  building its first window (an Office recursion; same on GE-Proton 11-6 and 11-7), so the automatic
  value is capped at 180, which renders and takes input correctly on GE-Proton11-7. GE-Proton11-6's
  driver mixed pixel and logical coordinates in the window geometry (input offset, no input at all
  at 168/180 dpi) and could get disconnected for committing a surface before its configure; the
  flake pins 11-7 for its Wayland fixes.
* **Direct2D comes from a different Wine** (`d2d1-fix/`): nixpkgs' Wine 11.16 `d2d1.dll` runs on
  GE-Proton 11's Wine 11.0. It only depends on public DLL interfaces, but if a future Proton bumps its
  Wine past the fix the override becomes unnecessary; if a future nixpkgs Wine adds a dependency the
  Proton base lacks, the launcher's log will show `d2d1` failing to load.

## Expectations

This is the college try, not a guarantee. Things that historically break: Microsoft account sign-in
(WebView2/Edge based), OneNote, Teams, and anything touching WinRT `Windows.*` APIs. If the
installer dies early, pin an older build with `MS365_VERSION`, or switch runners.
