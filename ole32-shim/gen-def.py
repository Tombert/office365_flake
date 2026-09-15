"""Generate a forwarder-only ole32.def from a Wine builtin ole32.dll.

Every export the builtin has is re-exported: forwarders (mostly to combase) are kept as-is,
everything else forwards to ole32_wine.dll, which is a renamed copy of the builtin that the
install script places next to the shim. Missing modern entry points are then added on top.
"""
import sys, pefile

src, out = sys.argv[1], sys.argv[2]
p = pefile.PE(src)
lines = ["LIBRARY ole32.dll", "EXPORTS"]
names = set()
for e in p.DIRECTORY_ENTRY_EXPORT.symbols:
    if not e.name:
        continue
    n = e.name.decode()
    names.add(n)
    if e.forwarder:
        lines.append(f"  {n} = {e.forwarder.decode()}")
    else:
        lines.append(f"  {n} = ole32_wine.{n}")

# Entry points Office asks ole32 for that this Wine build's ole32 does not export (and whose combase
# counterpart is only a stub that aborts when called). These are implemented in shim.c.
EXTRA = ["CoRegisterActivationFilter"]
for n in EXTRA:
    if n not in names:
        lines.append(f"  {n}")

open(out, "w").write("\n".join(lines) + "\n")
print(f"{len(lines)-2} exports written to {out}")
