# Cross-compile ms365keeper.exe (see keeper.c) for x86_64 Windows using zig's bundled clang + mingw.
{ stdenvNoCC, zig }:
stdenvNoCC.mkDerivation {
  pname = "ms365keeper";
  version = "0.1";
  src = ./.;
  nativeBuildInputs = [ zig ];
  dontConfigure = true;
  buildPhase = ''
    runHook preBuild
    export ZIG_GLOBAL_CACHE_DIR=$TMPDIR/zig-cache
    export ZIG_LOCAL_CACHE_DIR=$TMPDIR/zig-local
    zig cc -target x86_64-windows-gnu -O2 -Wall -Wl,--subsystem,windows -o ms365keeper.exe keeper.c -lntdll
    runHook postBuild
  '';
  installPhase = ''
    runHook preInstall
    install -Dm644 ms365keeper.exe $out/lib/wine/x86_64-windows/ms365keeper.exe
    runHook postInstall
  '';
  meta.description = "Keeps a prefix's shared Wine session (and umu's launcher service in it) alive until the last Office app closes";
}
