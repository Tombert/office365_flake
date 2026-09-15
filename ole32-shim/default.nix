# Forwarder ole32.dll for a given Proton runner (see gen-def.py). Office's mso30win32client
# does GetProcAddress(ole32, "CoRegisterActivationFilter") and crashes when it gets NULL; the
# Wine in GE-Proton 11 / Soda 11 implements it in combase but never exported it from ole32.
{ stdenvNoCC, zig, python3, runner, runnerName }:
let py = python3.withPackages (ps: [ ps.pefile ]);
in stdenvNoCC.mkDerivation {
  pname = "ole32-shim-${runnerName}";
  version = "0.1";
  src = ./.;
  nativeBuildInputs = [ zig py ];
  dontConfigure = true;
  buildPhase = ''
    runHook preBuild
    export ZIG_GLOBAL_CACHE_DIR=$TMPDIR/zig-cache
    export ZIG_LOCAL_CACHE_DIR=$TMPDIR/zig-local
    python3 gen-def.py ${runner}/files/lib/wine/x86_64-windows/ole32.dll ole32.def
    zig cc -target x86_64-windows-gnu -shared -O2 -Wall -o ole32.dll shim.c ole32.def
    runHook postBuild
  '';
  installPhase = ''
    runHook preInstall
    install -Dm644 ole32.dll $out/lib/wine/x86_64-windows/ole32.dll
    install -Dm644 ole32.def $out/share/ole32.def
    runHook postInstall
  '';
}
