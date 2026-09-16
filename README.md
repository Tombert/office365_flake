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
| `MS365_D2D_SYNC` | `1` | `0` turns off the ole32 shim's GPU wait after every Direct2D `EndDraw` (see the ribbon row below) |
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

Status as of 2026-09-15 with GE-Proton 11-6 and Microsoft 365 Apps build 16.0.20326.20144:

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
| Word: `CoRegisterActivationFilter` missing from ole32 (mso30win32client dereferences NULL) | `ole32-shim/`: forwarder `ole32.dll` that re-exports the builtin (kept as `ole32_wine.dll`) and implements the function |
| Word: `SetFileShortNameW`, `FindPackagesByPackageFamily`, `SetThreadpoolTimerEx` not exported by Wine's kernel32; the loader binds them to aborting stubs | the ole32 shim registers a loader notification and rewrites those import slots in every module with benign replacements |
| Word: special user APCs (`QueueUserAPC2`) dispatched wrongly by this Wine, Office aborts | the shim hides `QueueUserAPC2` from `GetProcAddress`; Office uses its pre-20H1 path |
| Word: splash-screen thread crashes in combase during COM apartment teardown, Office's crash handler kills the app | Word is launched with `/q` |
| `mso.dll` and friends live in Office's VFS tree; the App-V redirection layer is unreliable under Wine | the VFS tree is mirrored into `Program Files` with symlinks |
| Office offers safe mode after every crash via a modal prompt | the launcher clears the Resiliency key before starting an app |
| Sign-in dies with 53u4r / 12009 after the password (or on the email page) | Wine's winhttp/wininet reject unimplemented option codes with 12009; the shim accepts them (winhttp 77/140, wininet 11) |
| Licensing dialog fails with E_NOINTERFACE | shim serves `ILanguageStatics` and `IJsonObjectStatics`, which Wine's WinRT factories lack |
| Word exits at start on the legacy licensing path | product set to vNext licensing mode (LicensingNext = 2), SCA off by default |
| Ribbon font/size boxes, Comments/Editing/Share buttons and the title-bar search field are solid grey blocks; a control only shows its text while hovered, icons vanish under the hover highlight | Wine 11.0's `d2d1` ignores the fill mode of geometry groups. Office draws each control's border as two nested rounded rectangles with even-odd fill (a ring); Wine fills the union, and the compositor stretches that slab over the text. `d2d1-fix/` ships `d2d1.dll` from nixpkgs' Wine 11.16 (fixed upstream in February 2026) with its builtin signature blanked so Proton loads it, registered as a native override. The shim's `EndDraw` GPU wait (`MS365_D2D_SYNC`) predates this finding and is kept as a safety net |
| "Missing proofing tools" banner although the dictionaries are installed | Office finds proofing engines through MSI component registrations the Click-to-Run integrator never wrote under Wine; `msi-components.py` rebuilds them from the package manifests |

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
a stub, so Word exits with code 64 instead. Buy or manage the subscription on the web, then sign in
again in Word.

Other WinRT gaps the ole32 shim fills for the licensing code: `Windows.Globalization.Language`
statics (`ILanguageStatics`) and `Windows.Data.Json.JsonObject` statics (`Parse`/`TryParse`, built
on Wine's `JsonValue`).

## Known problems

* **Proofing engine**: the "missing proofing tools" banner is gone (the MSI component registrations
  are in place and the speller host loads), but the English speller engine itself is not picked up
  yet, so spell checking stays off. Next on the list.
* **Direct2D comes from a different Wine** (`d2d1-fix/`): nixpkgs' Wine 11.16 `d2d1.dll` runs on
  GE-Proton 11's Wine 11.0. It only depends on public DLL interfaces, but if a future Proton bumps its
  Wine past the fix the override becomes unnecessary; if a future nixpkgs Wine adds a dependency the
  Proton base lacks, the launcher's log will show `d2d1` failing to load.

## Expectations

This is the college try, not a guarantee. Things that historically break: Microsoft account sign-in
(WebView2/Edge based), OneNote, Teams, and anything touching WinRT `Windows.*` APIs. If the
installer dies early, pin an older build with `MS365_VERSION`, or switch runners.
