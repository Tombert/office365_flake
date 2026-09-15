{ lib
, stdenvNoCC
, writeShellApplication
, symlinkJoin
, makeDesktopItem
, umu-launcher
, curl
, coreutils
, gnused
, protonGE
, protosoda
, sppcShim
, ole32ShimGE
, ole32ShimSoda
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

  cli = writeShellApplication {
    name = "ms365";
    runtimeInputs = [ umu-launcher curl coreutils gnused ];
    text = ''
      MS365_PROTON_GE=${protonGE}
      MS365_PROTOSODA=${protosoda}
      MS365_SPPC_SHIM=${sppcShim}/lib/wine/x86_64-windows/sppc.dll
      MS365_OLE32_SHIM_GE=${ole32ShimGE}/lib/wine/x86_64-windows/ole32.dll
      MS365_OLE32_SHIM_SODA=${ole32ShimSoda}/lib/wine/x86_64-windows/ole32.dll
      export MS365_PROTON_GE MS365_PROTOSODA MS365_SPPC_SHIM MS365_OLE32_SHIM_GE MS365_OLE32_SHIM_SODA
      ${builtins.readFile ./ms365.sh}
    '';
  };

  appWrappers = lib.mapAttrsToList (key: app: writeShellApplication {
    name = "ms365-${key}";
    runtimeInputs = [ cli ];
    text = ''exec ms365 run ${key} "$@"'';
  }) apps;

  desktopItems = lib.mapAttrsToList (key: app: makeDesktopItem {
    name = "ms365-${key}";
    desktopName = "${app.name} (Proton)";
    genericName = app.name;
    exec = "ms365-${key} %F";
    icon = "ms365-${key}";
    terminal = false;
    categories = [ "Office" ];
    mimeTypes = lib.filter (s: s != "") (lib.splitString ";" app.mime);
    startupNotify = true;
  }) apps;
in
symlinkJoin {
  name = "ms365";
  paths = [ cli ] ++ appWrappers ++ desktopItems;
  meta = {
    description = "Run Microsoft 365 click-to-run Office through umu-launcher and GE-Proton";
    mainProgram = "ms365";
    platforms = [ "x86_64-linux" ];
  };
}
