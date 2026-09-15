# Cross-compile the sppc.dll shim (see sppc.c) for x86_64 Windows using zig's bundled clang + mingw headers.
{ stdenvNoCC, zig }:
stdenvNoCC.mkDerivation {
  pname = "sppc-shim";
  version = "0.1";
  src = ./.;
  nativeBuildInputs = [ zig ];
  dontConfigure = true;
  buildPhase = ''
    runHook preBuild
    export ZIG_GLOBAL_CACHE_DIR=$TMPDIR/zig-cache
    export ZIG_LOCAL_CACHE_DIR=$TMPDIR/zig-local
    zig cc -target x86_64-windows-gnu -shared -O2 -Wall -o sppc.dll sppc.c
    runHook postBuild
  '';
  installPhase = ''
    runHook preInstall
    install -Dm644 sppc.dll $out/lib/wine/x86_64-windows/sppc.dll
    runHook postInstall
  '';
  meta.description = "Benign sppc.dll (Software Protection Platform) replacement so Office click-to-run can install licences under Wine";
}
