# Direct2D from a newer Wine, for the Wine 11.0 base of GE-Proton 11 / Soda 11.
#
# Office draws every ribbon control into a sprite atlas with Direct2D. Boxes such as the font name
# and size combos, the Comments/Editing/Share buttons and the title-bar search field get their border
# as a geometry group of two rounded rectangles filled with D2D1_FILL_MODE_ALTERNATE (even-odd), i.e.
# a ring. Wine 11.0's d2d1 ignores the group's fill mode ("d2d_geometry_group_init Ignoring
# fill_mode"), fills the union instead, and the "border" becomes a solid slab that the compositor
# stretches over the control's text: grey blocks that only show text while hovered. Wine fixed
# geometry groups in February 2026 ("d2d1: Create a path internally for the geometry group"), so
# take d2d1.dll from a Wine that has it. d2d1 only imports public APIs (d3d10_1, d3dcompiler, gdi32,
# user32, ntdll, ucrtbase), so it runs fine on the older Proton Wine.
#
# Wine loads a "Wine builtin DLL" found in system32 by looking the name up in its own DLL directories
# instead, so the copy would be ignored; blank the signature (offset 64 of the DOS stub) and let the
# launcher register the file as a native override.
{ stdenvNoCC, wine }:
stdenvNoCC.mkDerivation {
  pname = "d2d1-fix";
  version = wine.version;
  dontUnpack = true;
  buildPhase = ''
    runHook preBuild
    cp ${wine}/lib/wine/x86_64-windows/d2d1.dll d2d1.dll
    chmod u+w d2d1.dll
    sig="Wine builtin DLL"
    [ "$(dd if=d2d1.dll bs=1 skip=64 count=16 2>/dev/null)" = "$sig" ] || { echo "no builtin signature at offset 64"; exit 1; }
    printf 'ms365 native DLL' | dd of=d2d1.dll bs=1 seek=64 count=16 conv=notrunc 2>/dev/null
    runHook postBuild
  '';
  installPhase = ''
    runHook preInstall
    install -Dm644 d2d1.dll $out/lib/wine/x86_64-windows/d2d1.dll
    runHook postInstall
  '';
  meta.description = "Wine ${wine.version} d2d1.dll with the geometry-group fill-mode fix, loadable as a native DLL";
}
