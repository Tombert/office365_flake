# In-process COM server providing CLSID_CUIAutomationRegistrar (see uia.c), which Wine lacks and
# Office requires once a document window has focus.
{ stdenvNoCC, zig }:
stdenvNoCC.mkDerivation {
  pname = "ms365uia";
  version = "0.1";
  src = ./.;
  nativeBuildInputs = [ zig ];
  dontConfigure = true;
  buildPhase = ''
    runHook preBuild
    export ZIG_GLOBAL_CACHE_DIR=$TMPDIR/zig-cache
    export ZIG_LOCAL_CACHE_DIR=$TMPDIR/zig-local
    zig cc -target x86_64-windows-gnu -shared -O2 -Wall -o ms365uia.dll uia.c -lole32 -luuid
    runHook postBuild
  '';
  installPhase = ''
    runHook preInstall
    install -Dm644 ms365uia.dll $out/lib/wine/x86_64-windows/ms365uia.dll
    runHook postInstall
  '';
}
