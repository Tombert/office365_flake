{ lib
, stdenvNoCC
, writeShellApplication
, symlinkJoin
, makeDesktopItem
, umu-launcher
, curl
, coreutils
, gnused
, python3
, protonGE
, protosoda
, sppcShim
, ole32ShimGE
, ole32ShimSoda
, uiaShim
, d2d1Fix
, selawik
  # Variant knobs: the same launcher, packaged under another command name with other defaults
  # (e.g. a retail Office 2024 SKU in its own prefix next to Microsoft 365).
, cliName ? "ms365"
, description ? "Run Microsoft 365 click-to-run Office through umu-launcher and GE-Proton"
, desktopSuffix ? " (Proton)"     # appended to each .desktop entry's name
, only ? null                     # null = every app below; else a list of app keys to ship wrappers for
, defaults ? { }                  # MS365_* defaults baked in; the environment still overrides them
}:

let
  apps = {
    word       = { exe = "WINWORD.EXE";  name = "Microsoft Word";       mime = "application/msword;application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/rtf;"; };
    excel      = { exe = "EXCEL.EXE";    name = "Microsoft Excel";      mime = "application/vnd.ms-excel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;text/csv;"; };
    powerpoint = { exe = "POWERPNT.EXE"; name = "Microsoft PowerPoint"; mime = "application/vnd.ms-powerpoint;application/vnd.openxmlformats-officedocument.presentationml.presentation;"; };
    outlook    = { exe = "OUTLOOK.EXE";  name = "Microsoft Outlook";    mime = "message/rfc822;"; };
    onenote    = { exe = "ONENOTE.EXE";  name = "Microsoft OneNote";    mime = ""; };
    access     = { exe = "MSACCESS.EXE"; name = "Microsoft Access";     mime = "application/msaccess;"; };
    publisher  = { exe = "MSPUB.EXE";    name = "Microsoft Publisher";  mime = "application/vnd.ms-publisher;"; };
  };

  shippedApps = if only == null then apps else lib.getAttrs only apps;

  defaultsSh = lib.concatStrings (lib.mapAttrsToList (k: v: ": \"\${${k}:=${v}}\"\n") defaults);

  cli = writeShellApplication {
    name = cliName;
    runtimeInputs = [ umu-launcher curl coreutils gnused python3 ];
    text = ''
      : "''${MS365_CLI:=${cliName}}"
      ${defaultsSh}
      MS365_PROTON_GE=${protonGE}
      MS365_PROTOSODA=${protosoda}
      MS365_SPPC_SHIM=${sppcShim}/lib/wine/x86_64-windows/sppc.dll
      MS365_OLE32_SHIM_GE=${ole32ShimGE}/lib/wine/x86_64-windows/ole32.dll
      MS365_OLE32_SHIM_SODA=${ole32ShimSoda}/lib/wine/x86_64-windows/ole32.dll
      MS365_UIA_SHIM=${uiaShim}/lib/wine/x86_64-windows/ms365uia.dll
      MS365_D2D1_DLL=${d2d1Fix}/lib/wine/x86_64-windows/d2d1.dll
      MS365_MSI_COMPONENTS=${./msi-components.py}
      MS365_UI_FONTS=${selawik}
      export MS365_UI_FONTS MS365_PROTON_GE MS365_PROTOSODA MS365_SPPC_SHIM MS365_OLE32_SHIM_GE MS365_OLE32_SHIM_SODA MS365_UIA_SHIM MS365_D2D1_DLL MS365_MSI_COMPONENTS
      ${builtins.readFile ./ms365.sh}
    '';
  };

  appWrappers = lib.mapAttrsToList (key: app: writeShellApplication {
    name = "${cliName}-${key}";
    runtimeInputs = [ cli ];
    text = ''exec ${cliName} run ${key} "$@"'';
  }) shippedApps;

  desktopItems = lib.mapAttrsToList (key: app: makeDesktopItem {
    name = "${cliName}-${key}";
    desktopName = "${app.name}${desktopSuffix}";
    genericName = app.name;
    exec = "${cliName}-${key} %F";
    icon = "${cliName}-${key}";
    terminal = false;
    categories = [ "Office" ];
    mimeTypes = lib.filter (s: s != "") (lib.splitString ";" app.mime);
    startupNotify = true;
  }) shippedApps;
in
symlinkJoin {
  name = cliName;
  paths = [ cli ] ++ appWrappers ++ desktopItems;
  meta = {
    inherit description;
    mainProgram = cliName;
    platforms = [ "x86_64-linux" ];
  };
}
