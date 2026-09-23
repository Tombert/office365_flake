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
* Products: UserData\...\Products\<squished product>\InstallProperties with WindowsInstaller=1 (the
  install-state check in MsiGetComponentPath) and Software\Classes\Installer\Products\<squished
  product> (MsiGetProductCode validates every product that registered a component against it).
* Features (MsiQueryFeatureState, which Office runs on e.g. OfficeMSProof6 before it trusts the
  proofing host): Software\Classes\Installer\Features\<product> naming the parent feature, and
  UserData\...\Products\<product>\Features holding the feature's components in base85.
* Qualified components (MsiEnumComponentQualifiers / MsiProvideQualifiedComponent): the
  PublishComponent entries. Office enumerates e.g. the speller category to learn which languages
  have proofing tools, then asks for the file behind a category and language. The value lives under
  Software\Classes\Installer\Components\<squished category id>, named by the qualifier, as a
  REG_MULTI_SZ descriptor "<product code base85><feature>><component base85><app data>".
  The manifests do not say which of a feature's components a category refers to; ROLE picks it by
  file role (speller category -> MSSP*.LEX, ProofDataFile -> CSS7DATA*.DLL, ...), or else by the
  file the qualifier names (VBA's "vbe.dll_7.1" -> VBE7.DLL, "1033\solver.xlam"). The rest get the
  feature's first file component.
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
    "{EF8E9806-D488-4BE1-8D06-01B401C9DE98}": r"MSSP\d*\w*\.LEX$",   # "<lang>\Normal": the language's normal spelling dictionary
    "{99A98A3E-2336-44D4-B56A-099F5CE3AF98}": r"MSWDS_\w*\.LEX$",    # find all word forms
    "{7228F250-A792-4094-A843-2EAAB88516C8}": r"FILTER_PHRASE\.BIN$",
    "{48960E6E-4B74-4873-9503-2D76FECCAA90}": r"\\\.CONFIG$",
}


def qualifier_file(qualifier):
    """The file a qualifier names, normalised: "1033\\solver.xlam" -> "solver.xlam", "vbe.dll_7.1" ->
    "vbe.dll" (VBA's version suffix), "stintl.dll\\1033" -> "stintl.dll", "x.x86.dll" -> "x.dll"."""
    for part in qualifier.split("\\"):
        m = re.match(r"^([A-Za-z][\w .-]*?\.[A-Za-z]\w{0,3})(?:_[\w.]+)?$", part)
        if m:
            return m.group(1).lower().replace(".x86", "").replace(" ", "")
    return None


def by_file_name(name, candidates, comps):
    """The component whose key path file is `name`, or `name` with a number before the extension
    (vbe.dll -> VBE7.DLL, vbeext.olb -> VBEEXT1.OLB, mspst.dll -> MSPST32.DLL)."""
    stem, ext = os.path.splitext(name)
    numbered = re.compile(re.escape(stem) + r"\d+" + re.escape(ext) + "$")
    files = [(c, comps[c].rsplit("\\", 1)[-1].lower().replace(" ", "")) for c in candidates
             if resolve(comps.get(c, "")) and not comps[c].endswith("\\")]
    for match in (lambda f: f == name, numbered.match):
        for c, f in files:
            if match(f):
                return c
    return None


def choose_component(category, qualifier, feature_comps, comps):
    pat = ROLE.get(category.upper())
    if pat:
        for c in feature_comps:
            if re.search(pat, comps.get(c, ""), re.I):
                return c
    # Most other qualifiers name the file (VBA's "vbe.dll_7.1", an add-in's "1033\\solver.xlam", a
    # sound's "arrow.wav"); Office opens whatever the descriptor resolves to, so the first file of
    # the feature (msvcr100.dll for VBAFiles: "cannot access" every add-in with a VBA project) is
    # wrong for them.
    name = qualifier_file(qualifier)
    c = name and by_file_name(name, feature_comps, comps)
    if c:
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


attr_re = re.compile(r'([A-Za-z]+)="([^"]*)"')
feat_re = re.compile(r'<Feature ([^>]*)>(.*?)</Feature>', re.S)


def elements(text, tag):
    """Attribute dicts of every <tag ...> element, whatever the attribute order."""
    for m in re.finditer(r'<%s\s([^>]*?)/?>' % tag, text):
        yield {k: html.unescape(v) for k, v in attr_re.findall(m.group(1))}


def feature_tree(text):
    """feature id -> (parent id, [component ids])"""
    tree = {}
    for attrs, body in feat_re.findall(text):
        a = dict(attr_re.findall(attrs))
        tree[a.get("FeatureId", "")] = (a.get("Parent", ""), ["{%s}" % c.upper() for c in re.findall(r'ComponentId="\{([0-9A-Fa-f-]+)\}"', body)])
    return tree

components = {}   # component id -> {product codes} ; path in comp_path
comp_path = {}
reg_keys = set()
products = {pkg_id.upper(): "Microsoft 365 (Click-to-Run package)"}
features = {}     # product -> feature -> (parent, [component ids])
qualified = {}    # category -> {qualifier: [descriptor strings]}

for manifest in sorted(glob.glob(os.path.join(pkg_dir, "C2RManifest.*.xml"))):
    text = read_manifest(manifest)
    m = re.search(r'ProductCode="\{([0-9A-Fa-f-]+)\}"', text)
    product = m.group(1).upper() if m else pkg_id.upper()
    products.setdefault(product, os.path.basename(manifest)[len("C2RManifest."):-len(".xml")])
    comps = {}
    for a in elements(text, "Component"):
        if "KeyPath" not in a:
            continue
        comp, keypath = a["ComponentId"].upper(), a["KeyPath"]
        comps[comp] = keypath
        path = resolve(keypath)
        if path:
            comp_path.setdefault(comp.strip("{}"), path)
            components.setdefault(comp.strip("{}"), set()).update({pkg_id.upper(), product})
        else:
            rk = reg_keypath(keypath)
            if rk:
                reg_keys.add(rk)
    tree = feature_tree(text)
    for fid, (parent, clist) in tree.items():
        clist = [c for c in clist if c.strip("{}") in comp_path]
        features.setdefault(product, {})[fid] = (parent, clist)
        merged = features.setdefault(pkg_id.upper(), {}).setdefault(fid, (parent, []))
        merged[1].extend(c for c in clist if c not in merged[1])
    for a in elements(text, "PublishComponent"):
        category, qualifier, appdata, feature = a["PublishComponentId"], a.get("Qualifier", ""), a.get("AppData", ""), a.get("Feature", "")
        comp = a["ComponentId"].upper() if a.get("ComponentId") else choose_component(category, qualifier, tree.get(feature, ("", []))[1], comps)
        if not comp or comp.strip("{}") not in comp_path:
            continue
        desc = base85(product) + feature + ">" + base85(comp) + appdata
        qualified.setdefault(category.upper(), {}).setdefault(qualifier, []).append(desc)

base = r"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18"
lines = ["Windows Registry Editor Version 5.00", ""]
for prod, name in sorted(products.items()):
    lines += ["[%s\\Products\\%s\\InstallProperties]" % (base, squish(prod)),
              '"WindowsInstaller"=dword:00000001',
              '"InstallLocation"="%s"' % office_root.replace("\\", "\\\\"),
              '"DisplayName"="%s"' % name, "",
              r"[HKEY_LOCAL_MACHINE\Software\Classes\Installer\Products\%s]" % squish(prod),
              '"ProductName"="%s"' % name, ""]
    feats = features.get(prod, {})
    if feats:
        lines.append(r"[HKEY_LOCAL_MACHINE\Software\Classes\Installer\Features\%s]" % squish(prod))
        lines += ['"%s"="%s"' % (reg_name(f), reg_name(parent)) for f, (parent, _) in sorted(feats.items())]
        lines += ["", "[%s\\Products\\%s\\Features]" % (base, squish(prod))]
        for f, (parent, clist) in sorted(feats.items()):
            data = "".join(base85(c) for c in clist) + ("\x02" + parent if parent else "")
            lines.append('"%s"=%s' % (reg_name(f), ('"%s"' % reg_name(data)) if "\x02" not in data else "hex(1):" + ",".join("%02x" % b for b in (data + "\0").encode("latin-1"))))
        lines.append("")
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
print("%d components, %d products, %d features, %d qualified-component categories, %d registry keys for package {%s} -> %s"
      % (len(components), len(products), sum(len(f) for f in features.values()), len(qualified), len(reg_keys), pkg_id, out))
