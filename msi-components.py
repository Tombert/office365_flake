#!/usr/bin/env python3
r"""Register Office's Click-to-Run MSI data in the Wine prefix's registry.

Office locates optional payload (proofing tools, converters, add-ins, ...) through the Windows
Installer API. On Windows the Click-to-Run integrator writes the registrations into its virtual
registry; under Wine that step produces an empty hive, so Word reports "missing proofing tools"
although every file is on disk. The per-package manifests that Click-to-Run leaves under ProgramData
carry everything the lookups need, so this script turns them into a .reg file for Wine's msi.dll:

* Components (MsiGetComponentPath): UserData\S-1-5-18\Components\<squished component id>, one value
  per product code (Office passes the App-V package id as the product, the descriptors below carry
  the package's own ProductCode, so both are registered) holding the component's key path.
* Products (the install-state check in MsiGetComponentPath): UserData\...\Products\<squished
  product>\InstallProperties with WindowsInstaller=1.
* Qualified components (MsiEnumComponentQualifiers / MsiProvideQualifiedComponent): the
  PublishComponent entries. Office enumerates e.g. the speller category to learn which languages
  have proofing tools, then asks for the file behind a category and language. The value lives under
  Software\Classes\Installer\Components\<squished category id>, named by the qualifier, as a
  REG_MULTI_SZ descriptor "<product code base85><feature>><component base85><app data>".
  The manifests do not say which of a feature's components a category refers to; ROLE picks it by
  file role (speller category -> MSSP*.LEX, ProofDataFile -> CSS7DATA*.DLL, ...). Unmatched
  categories get the feature's first file component.
* Registry key paths ("22:\Software\...", e.g. Common\InstalledPackages\<product>) are created as
  empty keys; Office checks some of them itself.

usage: msi-components.py <prefix drive_c> <output.reg>
"""
import glob
import html
import os
import re
import struct
import sys
import uuid

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
    return "".join(p[::-1] for p in parts).upper()


B85 = "!$%&'()*+,-.0123456789=?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[]^_`abcdefghijklmnopqrstuvwxyz{}~"


def base85(guid):
    """Windows Installer's packed GUID form: four little-endian DWORDs, five chars each, low digit first."""
    raw = uuid.UUID(guid.strip("{}")).bytes_le
    s = ""
    for (x,) in struct.iter_unpack("<I", raw):
        for _ in range(5):
            s += B85[x % 85]
            x //= 85
    return s


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

# MSI registry key paths: "<root><flags>:\path", root 0 HKCR, 1 HKCU, 2 HKLM, 3 HKU; flag 2 = 64-bit view.
REG_ROOTS = {"0": "HKEY_CLASSES_ROOT", "1": "HKEY_CURRENT_USER", "2": "HKEY_LOCAL_MACHINE", "3": "HKEY_USERS"}


def resolve(keypath):
    for var, real in VFS:
        if keypath.upper().startswith(var.upper()):
            return real + keypath[len(var):]
    return None  # registry / assembly key paths: not a file


def reg_keypath(keypath):
    m = re.match(r"^(\d)(\d):\\(.*)$", keypath)
    if not m:
        return None
    return REG_ROOTS[m.group(2)] + "\\" + m.group(3).rstrip("\\")


# Which component of a feature a qualified-component category stands for, by category id and a
# regex on the key path's file name. Office's proofing categories (speller, grammar, hyphenation,
# thesaurus, contextual speller data, autocorrect list, ...). Order matters only within a category.
ROLE = {
    "{C9A7733D-5B0C-4B9F-B42E-8143FC77F881}": r"MSSP\d*\w*\.LEX$",   # speller lexicon
    "{F7BCB4C0-3189-4A81-8C29-B145871CAEAF}": r"MSGR\d*\w*\.LEX$",   # grammar lexicon
    "{58CB715A-C494-4EA3-8465-013B366CBED6}": r"MSHY\d*\w*\.DLL$",   # hyphenation engine
    "{59EB536B-8331-417F-93F5-0145E2FD0AE4}": r"MSHY\d*\w*\.LEX$",   # hyphenation lexicon
    "{5788A17B-DD37-49C0-8E75-01CC754381AC}": r"\\PROOF\\$",          # hyphenation folder
    "{98B57FD8-C181-4111-8072-01D39ADBB739}": r"MSTH\d*\w*\.DLL$",   # thesaurus engine
    "{1B2EE6D9-FB0D-4D1C-90F2-017FFF471A97}": r"MSTH\d*\w*\.LEX$",   # thesaurus lexicon
    "{13462384-2D94-43AD-BA1E-687A27298820}": r"CSS7DATA\w*\.DLL$",  # ProofDataFile
    "{539009AB-E05A-41F3-B6A5-0D0A68AF8B24}": r"NL7MODELS\w*\.DLL$", # ProofModelFile
    "{509B6F9B-C120-4FFC-86A1-012E20A9D3BE}": r"MSCSS7[A-Z]{2}\.DLL$",  # nlg_CSS (contextual speller)
    "{B95B001F-4C25-4297-BCBF-1AE95C48D95E}": r"MSCSS7CM_\w*\.DUB$", # nlg_updates
    "{EF8E9806-D488-4BE1-8D06-01B401C9DE98}": r"MSO\.ACL$",          # autocorrect list "<lang>\Normal"
    "{99A98A3E-2336-44D4-B56A-099F5CE3AF98}": r"MSWDS_\w*\.LEX$",    # find all word forms
    "{7228F250-A792-4094-A843-2EAAB88516C8}": r"FILTER_PHRASE\.BIN$",
    "{48960E6E-4B74-4873-9503-2D76FECCAA90}": r"\\\.CONFIG$",
}


def choose_component(category, feature_comps, comps):
    pat = ROLE.get(category.upper())
    if pat:
        for c in feature_comps:
            if re.search(pat, comps.get(c, ""), re.I):
                return c
    files = [c for c in feature_comps if resolve(comps.get(c, "")) and not comps[c].endswith("\\")]
    return (files or feature_comps or [None])[0]


def read_manifest(path):
    with open(path, "rb") as f:
        data = f.read()
    if data.startswith(b"\xff\xfe"):
        return data.decode("utf-16")
    return data.decode("utf-8", "replace")


def reg_name(s):
    return s.replace("\\", "\\\\").replace('"', '\\"')


def multi_sz(strings):
    # regedit widens hex(7) data of an ANSI .reg file to UTF-16 itself, so one byte per character
    data = "".join(s + "\0" for s in strings) + "\0"
    hexbytes = ["%02x" % b for b in data.encode("latin-1")]
    # regedit line continuation: keep lines short
    chunks = [",".join(hexbytes[i:i + 24]) for i in range(0, len(hexbytes), 24)]
    return "hex(7):" + ",\\\n  ".join(chunks)


comp_re = re.compile(r'<Component\s+ComponentId="\{([0-9A-Fa-f-]+)\}"\s+KeyPath="([^"]*)"')
feat_re = re.compile(r'<Feature FeatureId="([^"]+)"[^>]*>(.*?)</Feature>', re.S)
pub_re = re.compile(r'<PublishComponent PublishComponentId="([^"]+)" Qualifier="([^"]*)" AppData="([^"]*)" Feature="([^"]*)"')

components = {}   # component id -> {product codes} ; path in comp_path
comp_path = {}
reg_keys = set()
products = {pkg_id.upper()}
qualified = {}    # category -> {qualifier: [descriptor strings]}

for manifest in sorted(glob.glob(os.path.join(pkg_dir, "C2RManifest.*.xml"))):
    text = read_manifest(manifest)
    m = re.search(r'ProductCode="\{([0-9A-Fa-f-]+)\}"', text)
    product = m.group(1).upper() if m else pkg_id.upper()
    products.add(product)
    comps = {}
    for comp, keypath in comp_re.findall(text):
        keypath = html.unescape(keypath)
        comps["{%s}" % comp.upper()] = keypath
        path = resolve(keypath)
        if path:
            comp_path.setdefault(comp.upper(), path)
            components.setdefault(comp.upper(), set()).update({pkg_id.upper(), product})
        else:
            rk = reg_keypath(keypath)
            if rk:
                reg_keys.add(rk)
    features = {}
    for fid, body in feat_re.findall(text):
        features[fid] = ["{%s}" % c.upper() for c in re.findall(r'ComponentId="\{([0-9A-Fa-f-]+)\}"', body)]
    for category, qualifier, appdata, feature in pub_re.findall(text):
        qualifier, appdata = html.unescape(qualifier), html.unescape(appdata)
        comp = choose_component(category, features.get(feature, []), comps)
        if not comp or comp.strip("{}") not in comp_path:
            continue
        desc = base85(product) + feature + ">" + base85(comp) + appdata
        qualified.setdefault(category.upper(), {}).setdefault(qualifier, []).append(desc)

base = r"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18"
lines = ["Windows Registry Editor Version 5.00", ""]
for prod in sorted(products):
    lines += ["[%s\\Products\\%s\\InstallProperties]" % (base, squish(prod)),
              '"WindowsInstaller"=dword:00000001',
              '"InstallLocation"="%s"' % office_root.replace("\\", "\\\\"), ""]
for comp, prods in sorted(components.items()):
    lines.append("[%s\\Components\\%s]" % (base, squish(comp)))
    for prod in sorted(prods):
        lines.append('"%s"="%s"' % (squish(prod), comp_path[comp].replace("\\", "\\\\")))
    lines.append("")
for category, quals in sorted(qualified.items()):
    lines.append(r"[HKEY_LOCAL_MACHINE\Software\Classes\Installer\Components\%s]" % squish(category))
    for qualifier, descs in sorted(quals.items()):
        lines.append('"%s"=%s' % (reg_name(qualifier), multi_sz(descs)))
    lines.append("")
for rk in sorted(reg_keys):
    lines += ["[%s]" % rk, ""]
with open(out, "w", newline="\r\n") as f:
    f.write("\n".join(lines))
print("%d components, %d products, %d qualified-component categories, %d registry keys for package {%s} -> %s"
      % (len(components), len(products), len(qualified), len(reg_keys), pkg_id, out))
