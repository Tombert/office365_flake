{
  description = "Microsoft 365 (Office click-to-run) on Linux via umu-launcher + GE-Proton / ProtoSoda";

  inputs.nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";

  outputs = { self, nixpkgs }:
    let
      system = "x86_64-linux";
      pkgs = import nixpkgs {
        inherit system;
        config.allowUnfree = true; # proton-ge-bin is unfree-tagged (Valve/MS redistributables)
      };
      lib = pkgs.lib;

      protonGE = pkgs.proton-ge-bin.steamcompattool;

      # Bottles' Soda Wine core wrapped in the Proton layout that umu expects.
      # Soda 11.0-8 is the build family in which the Bottles maintainer reported
      # Microsoft 365 installing, signing in (with 2FA) and running Word.
      protosoda = pkgs.stdenvNoCC.mkDerivation rec {
        pname = "protosoda";
        version = "11.0-3";
        src = pkgs.fetchurl {
          url = "https://github.com/bottlesdevs/wine/releases/download/protosoda-${version}/ProtoSoda-${version}.tar.gz";
          hash = "sha256-dAcQ50senPPLt/UhENC1Lq48vyTGh/ml9XnqRcDuauE=";
        };
        sourceRoot = ".";
        dontConfigure = true;
        dontBuild = true;
        dontFixup = true; # binaries run inside the Steam Linux Runtime container, leave them alone
        installPhase = ''
          runHook preInstall
          top=$(find . -mindepth 1 -maxdepth 1 -type d | head -n1)
          mkdir -p $out
          cp -a "$top"/. $out/
          runHook postInstall
        '';
        meta = with lib; {
          description = "Soda Wine core in Proton layout for umu (Bottles project)";
          homepage = "https://github.com/bottlesdevs/wine";
          license = licenses.lgpl21Plus;
          platforms = [ "x86_64-linux" ];
        };
      };

      sppcShim = pkgs.callPackage ./sppc { };
      uiaShim = pkgs.callPackage ./uia-shim { };
      ole32ShimGE = pkgs.callPackage ./ole32-shim { runner = protonGE; runnerName = "ge"; };
      ole32ShimSoda = pkgs.callPackage ./ole32-shim { runner = protosoda; runnerName = "protosoda"; };
      # newer Direct2D for both runners' Wine 11.0 base (ribbon controls rendered as grey blocks)
      d2d1Fix = pkgs.callPackage ./d2d1-fix { wine = pkgs.wine64Packages.unstable; };

      ms365 = pkgs.callPackage ./ms365.nix {
        inherit protonGE protosoda sppcShim ole32ShimGE ole32ShimSoda uiaShim d2d1Fix;
      };

      mkApp = exe: { type = "app"; program = "${ms365}/bin/${exe}"; };
    in
    {
      packages.${system} = {
        default = ms365;
        inherit ms365 protosoda protonGE sppcShim ole32ShimGE ole32ShimSoda uiaShim d2d1Fix;
      };

      apps.${system} = {
        default = mkApp "ms365";
        ms365 = mkApp "ms365";
        word = mkApp "ms365-word";
        excel = mkApp "ms365-excel";
        powerpoint = mkApp "ms365-powerpoint";
        outlook = mkApp "ms365-outlook";
        onenote = mkApp "ms365-onenote";
      };

      devShells.${system}.default = pkgs.mkShell {
        packages = [ ms365 pkgs.umu-launcher ];
      };
    };
}
