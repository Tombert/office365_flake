#!/usr/bin/env python3
"""Register Office's Click-to-Run MSI components in the Wine prefix's registry.

Office apps locate optional payload (proofing tools, converters, add-ins, ...) with
MsiGetComponentPath(package id, component id).  On Windows the Click-to-Run integrator
writes those component registrations; under Wine that step produces nothing (the App-V
registry hives in root/vreg stay empty), so Word reports e.g. "missing proofing tools"
although every file is on disk.  The per-package MSI manifests that Click-to-Run leaves
under ProgramData carry component ids and key paths, which is all MsiGetComponentPath
needs, so this script turns them into a .reg file for the per-machine installer data
that Wine's msi.dll consults.

usage: msi-components.py <prefix drive_c> <output.reg>
"""
import glob
import html
import os
import re
import sys

drive_c, out = sys.argv[1], sys.argv[2]
c2r = os.path.join(drive_c, "ProgramData", "Microsoft", "ClickToRun")
office_root = r"C:\Program Files\Microsoft Office\root"

# The App-V package id doubles as the MSI product code Office passes to MsiGetComponentPath.
pkg_dirs = [d for d in glob.glob(os.path.join(c2r, "{*}")) if os.path.isdir(d)]
if not pkg_dirs:
    sys.exit("no Click-to-Run package directory under " + c2r)
pkg_dir = pkg_dirs[0]
pkg_id = os.path.basename(pkg_dir).strip("{}")


def squish(guid):
    g = guid.strip("{}").replace("-", "")
    parts = [g[0:8], g[8:12], g[12:16]] + [g[16 + 2 * i:18 + 2 * i] for i in range(8)]
    return "".join(p[::-1] for p in parts)


# Click-to-Run's virtual file system: where each key-path root lives inside the prefix.
VFS = [
    ("%SFT_PROGRAM_FILES_X64%\\Microsoft Office\\", office_root + "\\"),
    ("%SFT_PROGRAM_FILES_X64%\\", office_root + "\\vfs\\ProgramFilesX64\\"),
    ("%SFT_PROGRAM_FILES_COMMON_X64%\\", office_root + "\\vfs\\ProgramFilesCommonX64\\"),
    ("%CSIDL_PROGRAM_FILES%\\Microsoft Office\\", office_root + "\\"),
    ("%CSIDL_PROGRAM_FILES%\\", office_root + "\\vfs\\ProgramFilesX86\\"),
    ("%CSIDL_PROGRAM_FILES_COMMON%\\", office_root + "\\vfs\\ProgramFilesCommonX86\\"),
    # %CSIDL_FONTS% (Office's bundled fonts) deliberately left out: registered, Office loads them
    # itself and the Fluent UI text (title bar, Share/Editing, message bars) turns into grey blocks.
    ("%SFT_SYSTEM32_X64%\\", office_root + "\\vfs\\System\\"),
    ("%CSIDL_SYSTEM%\\", office_root + "\\vfs\\SystemX86\\"),
    ("%CSIDL_WINDOWS%\\", office_root + "\\vfs\\Windows\\"),
    ("%CSIDL_COMMON_APPDATA%\\", office_root + "\\vfs\\Common AppData\\"),
]


def resolve(keypath):
    for var, real in VFS:
        if keypath.upper().startswith(var.upper()):
            return real + keypath[len(var):]
    return None  # registry / assembly key paths: not a file, skip


def read_manifest(path):
    with open(path, "rb") as f:
        data = f.read()
    if data.startswith(b"\xff\xfe"):
        return data.decode("utf-16")
    return data.decode("utf-8", "replace")


comp_re = re.compile(r'<Component\s+ComponentId="\{([0-9A-Fa-f-]+)\}"\s+KeyPath="([^"]*)"')
components = {}
for manifest in sorted(glob.glob(os.path.join(pkg_dir, "C2RManifest.*.xml"))):
    for comp, keypath in comp_re.findall(read_manifest(manifest)):
        path = resolve(html.unescape(keypath))
        if path:
            components.setdefault(comp.upper(), path)

prod = squish(pkg_id)
base = r"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18"
lines = ["Windows Registry Editor Version 5.00", "",
         "[%s\\Products\\%s\\InstallProperties]" % (base, prod),
         '"WindowsInstaller"=dword:00000001', '"InstallLocation"="%s"' % office_root.replace("\\", "\\\\"), ""]
for comp, path in sorted(components.items()):
    lines.append("[%s\\Components\\%s]" % (base, squish(comp)))
    lines.append('"%s"="%s"' % (prod, path.replace("\\", "\\\\")))
    lines.append("")
with open(out, "w", newline="\r\n") as f:
    f.write("\n".join(lines))
print("%d components for package {%s} -> %s" % (len(components), pkg_id, out))
