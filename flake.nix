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

      # nixpkgs' proton-ge-bin lags GE releases; GE-Proton11-7 (2026-09-16) is mostly Wine-Wayland
      # fixes (presentation geometry, configure completion, popup handling, toplevel sizing), which is
      # what MS365_WAYLAND=1 needs on HiDPI outputs. Override the version until nixpkgs catches up.
      protonGEPkg = pkgs.proton-ge-bin.overrideAttrs (final: prev:
        let
          release = pkgs.fetchzip {
            url = "https://github.com/GloriousEggroll/proton-ge-custom/releases/download/${final.version}/${final.version}-x86_64.tar.gz";
            hash = "sha256-ftW0vE45v2JsbaYqo/So0ZFfvdtakHX0XEXEE4TdxLk=";
          };
          # GE's Wayland driver stacks the subsurfaces of self-presenting child windows in the order
          # they were created or last moved, not in Win32 z-order (see wayland-fix/). Swap in the
          # patched winewayland.so. Wine finds its unix libraries next to the resolved ntdll.so, so
          # those (and the loaders in files/bin) are real copies; everything else stays a symlink.
          # The template prefix is copied as it is (real registry files, relative builtin-DLL
          # links): Proton copies it into every new prefix symlinks-as-symlinks, and store symlinks
          # there make system.reg and .update-timestamp read-only ("Read-only file system").
          src = pkgs.runCommand "${final.version}-x86_64-ms365" { } ''
            cp -rs ${release}/. $out
            chmod -R u+w $out
            for f in $out/files/lib/wine/x86_64-unix/* $out/files/bin/*; do
              [ -L "$f" ] && cp --remove-destination "$(readlink -f "$f")" "$f"
            done
            for d in $out/files/share/default_pfx*; do
              rm -rf "$d"
              cp -a ${release}/files/share/"$(basename "$d")" "$d"
            done
            install -m755 ${./wayland-fix/winewayland.so} $out/files/lib/wine/x86_64-unix/winewayland.so
          '';
        in {
          version = "GE-Proton11-7";
          inherit src;
          passthru = prev.passthru // {
            variants.x86_64-linux = { toolName = "${final.version}-x86_64"; inherit src; };
          };
        });
      protonGE = protonGEPkg.steamcompattool;

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

      # Selawik: Microsoft's open-source (MIT) metric-compatible stand-in for Segoe UI, the Office UI
      # font. Proton's prefix template maps Segoe UI to Times New Roman, which OneNote's canvas shows.
      selawik = pkgs.fetchzip {
        url = "https://github.com/microsoft/Selawik/releases/download/1.01/Selawik_Release.zip";
        hash = "sha256-BbjXJ8HFXrRklMOnGXyZIZeQ5Oksda4AqQXHmNqN6AQ=";
        stripRoot = false;
      };

      launcherDeps = { inherit protonGE protosoda sppcShim ole32ShimGE ole32ShimSoda uiaShim d2d1Fix selawik; };

      ms365 = pkgs.callPackage ./ms365.nix launcherDeps;

      # Office Home 2024, the retail one-time purchase: the same launcher with the retail SKU, in its
      # own prefix so it can sit next to a Microsoft 365 install. Retail 2024 builds come from the
      # Current channel. The product key is not given to the installer (Wine has no Software
      # Protection Platform to hold it); a key redeemed at setup.office.com is attached to the
      # Microsoft account, and Office licenses itself from that account after sign-in (vNext).
      # Home 2024 is Word, Excel, PowerPoint and OneNote.
      office2024 = pkgs.callPackage ./ms365.nix (launcherDeps // {
        cliName = "office2024";
        description = "Run Office Home 2024 (retail) through umu-launcher and GE-Proton";
        desktopSuffix = " 2024 (Proton)";
        only = [ "word" "excel" "powerpoint" "onenote" ];
        defaults = {
          MS365_HOME = "$HOME/.local/share/office2024";
          MS365_PRODUCT = "Home2024Retail";
          MS365_CHANNEL = "Current";
        };
      });

      # `nix run .#shortcuts`: menu entries, Desktop shortcuts, icons and file associations for an
      # installed Office (see shortcuts.sh).
      shortcuts = pkgs.writeShellApplication {
        name = "office-shortcuts";
        runtimeInputs = [ pkgs.coreutils pkgs.gnused pkgs.findutils ];
        text = ''
          SHORTCUTS_OFFICE2024=${office2024}
          SHORTCUTS_MS365=${ms365}
          SHORTCUTS_ICOUTILS=${pkgs.icoutils}
          SHORTCUTS_PYTHON=${pkgs.python3}
          SHORTCUTS_XDG_UTILS=${pkgs.xdg-utils}
          ${builtins.readFile ./shortcuts.sh}
        '';
      };

      mkApp = exe: { type = "app"; program = "${ms365}/bin/${exe}"; };
      mkApp2024 = exe: { type = "app"; program = "${office2024}/bin/${exe}"; };
    in
    {
      packages.${system} = {
        default = ms365;
        inherit ms365 office2024 shortcuts protosoda protonGE sppcShim ole32ShimGE ole32ShimSoda uiaShim d2d1Fix;
      };

      apps.${system} = {
        default = mkApp "ms365";
        ms365 = mkApp "ms365";
        word = mkApp "ms365-word";
        excel = mkApp "ms365-excel";
        powerpoint = mkApp "ms365-powerpoint";
        outlook = mkApp "ms365-outlook";
        onenote = mkApp "ms365-onenote";

        office2024 = mkApp2024 "office2024";
        word2024 = mkApp2024 "office2024-word";
        excel2024 = mkApp2024 "office2024-excel";
        powerpoint2024 = mkApp2024 "office2024-powerpoint";
        onenote2024 = mkApp2024 "office2024-onenote";

        shortcuts = { type = "app"; program = "${shortcuts}/bin/office-shortcuts"; };
      };

      devShells.${system}.default = pkgs.mkShell {
        packages = [ ms365 office2024 pkgs.umu-launcher ];
      };
    };
}
