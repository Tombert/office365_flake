/*
 * sppc.dll for running Microsoft Office click-to-run under Wine/Proton.
 *
 * Wine's own sppc.dll (Software Protection Platform client) is almost entirely stubs, and calling
 * a stub aborts the process. This replacement is a small, honest licence store:
 *
 *  - SLInstallLicense keeps the XrML licence files the Office integrator installs (under
 *    %ProgramData%\ms365-spp) and SLUninstallLicense removes them.
 *  - The query functions answer from those files: which product SKUs exist for an application,
 *    their names, the licence blobs.
 *  - Nothing is ever *granted*: there is no product key store and SLConsumeRight always reports
 *    that the right is not granted, so Office sees "installed, not activated" exactly like a fresh
 *    Windows install and goes to its own sign-in / activation flow (which lives outside SPP).
 *
 * x86_64 only: the Win64 ABI is caller-clean, so a wrong parameter count is harmless. Trace output
 * goes to OutputDebugString (visible with PROTON_LOG / WINEDEBUG=+debugstr).
 */
#include <windows.h>
#include <string.h>

typedef HANDLE HSLC;
typedef HANDLE HSLP;
typedef GUID   SLID;

/* values from Wine's include/slerror.h / slpublic.h */
#define SL_E_RIGHT_NOT_CONSUMED    ((HRESULT)0xC004F002)
#define SL_E_VALUE_NOT_FOUND       ((HRESULT)0xC004F012)
#define SL_E_RIGHT_NOT_GRANTED     ((HRESULT)0xC004F013)
#define SL_E_PKEY_NOT_INSTALLED    ((HRESULT)0xC004F014)

enum { SL_ID_APPLICATION = 0, SL_ID_PRODUCT_SKU = 1, SL_ID_LICENSE_FILE = 2, SL_ID_LICENSE = 3,
       SL_ID_PKEY = 4, SL_ID_ALL_LICENSES = 5, SL_ID_ALL_LICENSE_FILES = 6, SL_ID_STORE_TOKEN = 7 };
enum { SL_DATA_NONE = 0, SL_DATA_SZ = 1, SL_DATA_BINARY = 3, SL_DATA_DWORD = 4, SL_DATA_MULTI_SZ = 7 };
enum { SL_LICENSING_STATUS_UNLICENSED = 0, SL_LICENSING_STATUS_LICENSED = 1,
       SL_LICENSING_STATUS_IN_GRACE_PERIOD = 2, SL_LICENSING_STATUS_NOTIFICATION = 3 };

typedef struct {
    SLID     SkuId;
    int      eStatus;
    DWORD    dwGraceTime;
    DWORD    dwTotalGraceDays;
    HRESULT  hrReason;
    UINT64   qwValidityExpiration;
} SL_LICENSING_STATUS;

#define API __declspec(dllexport) HRESULT WINAPI

/* ---- tracing ------------------------------------------------------------ */
static void guid_str(const SLID *g, char *out)
{
    if (!g) { lstrcpyA(out, "(null)"); return; }
    wsprintfA(out, "{%08lx-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}", g->Data1, g->Data2, g->Data3,
              g->Data4[0], g->Data4[1], g->Data4[2], g->Data4[3], g->Data4[4], g->Data4[5], g->Data4[6], g->Data4[7]);
}
static void tracef(const char *fmt, ...)
{
    char buf[512];
    va_list ap; va_start(ap, fmt);
    wvsprintfA(buf, fmt, ap);
    va_end(ap);
    OutputDebugStringA(buf);
}
static void trace2(const char *fn, const SLID *a, const SLID *b, PCWSTR name, HRESULT hr)
{
    char ga[48], gb[48];
    guid_str(a, ga); guid_str(b, gb);
    tracef("sppc shim: %s id1=%s id2=%s name=%ls -> 0x%08lx", fn, ga, gb, name ? name : L"", hr);
}
#define TRACE_RET(fn, a, b, name, hr) do { HRESULT _hr = (hr); trace2(fn, a, b, name, _hr); return _hr; } while (0)

/* ---- licence store ------------------------------------------------------ */
#define MAX_LIC 512
#define MAX_IDS 8
struct policy { char *name; int type; char *value; };   /* type: SL_DATA_DWORD / SL_DATA_SZ / SL_DATA_BINARY (base64 text) */

struct lic {
    SLID   id;                 /* licenseId of the outer r:license */
    struct policy *pol; int npol;   /* every <sl:policyInt|Str|Bin> in the licence */
    char   idstr[40];
    char   title[128];
    char   family[128];        /* Security-SPP-Reserved-Family, e.g. Office16O365ProPlusR_Subscription1 */
    SLID   apps[MAX_IDS]; int napps;
    SLID   skus[MAX_IDS]; int nskus;
    DWORD  grace_days;
    char   path[MAX_PATH];
};
static struct lic store[MAX_LIC];
static int nstore = -1;
static char store_dir[MAX_PATH];
static CRITICAL_SECTION lock;

static int hexval(char c)
{
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return -1;
}
/* parse "{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}" (braces optional) */
static int parse_guid(const char *s, SLID *g)
{
    if (*s == '{') s++;
    BYTE b[16]; int bi = 0;
    for (int i = 0; i < 36; i++) {
        if (i == 8 || i == 13 || i == 18 || i == 23) { if (s[i] != '-') return 0; continue; }
        int hi = hexval(s[i]), lo = hexval(s[i + 1]);
        if (hi < 0 || lo < 0) return 0;
        b[bi++] = (BYTE)(hi * 16 + lo); i++;
    }
    g->Data1 = ((DWORD)b[0] << 24) | ((DWORD)b[1] << 16) | ((DWORD)b[2] << 8) | b[3];
    g->Data2 = (WORD)((b[4] << 8) | b[5]);
    g->Data3 = (WORD)((b[6] << 8) | b[7]);
    memcpy(g->Data4, b + 8, 8);
    return 1;
}
static int guid_eq(const SLID *a, const SLID *b) { return a && b && memcmp(a, b, sizeof(SLID)) == 0; }

static void add_id(SLID *arr, int *n, const SLID *g)
{
    for (int i = 0; i < *n; i++) if (guid_eq(&arr[i], g)) return;
    if (*n < MAX_IDS) arr[(*n)++] = *g;
}

static void copy_text(const char *start, const char *endtag, char *out, size_t cap)
{
    const char *e = strstr(start, endtag);
    size_t n = e ? (size_t)(e - start) : 0;
    if (n >= cap) n = cap - 1;
    memcpy(out, start, n); out[n] = 0;
}

/* Extract what we need from an XrML licence blob (UTF-8, possibly with BOM). */
static int parse_license(const char *xml, size_t len, struct lic *L)
{
    (void)len;
    memset(L, 0, sizeof(*L));
    const char *p = strstr(xml, "<r:license");
    if (!p) return 0;
    const char *q = strstr(p, "licenseId=\"");
    if (!q) return 0;
    q += 11;
    if (!parse_guid(q, &L->id)) return 0;
    guid_str(&L->id, L->idstr);
    if ((q = strstr(p, "<r:title>"))) copy_text(q + 9, "</r:title>", L->title, sizeof(L->title));
    if ((q = strstr(p, "name=\"Security-SPP-Reserved-Family\""))) {
        q = strchr(q, '>');
        if (q) copy_text(q + 1, "</sl:policyStr>", L->family, sizeof(L->family));
    }
    /* policy-definition licences (PPD) identify their SKU family through the editionId they apply to */
    if (!L->family[0] && (q = strstr(p, "<editionId"))) {
        const char *v = strstr(q, "value=\"");
        if (v && v - q < 200) copy_text(v + 7, "\"", L->family, sizeof(L->family));
    }
    /* collect every <sl:policyInt|Str|Bin name="..."[ attributes=...]>value</sl:policy...> */
    {
        int cap = 0;
        for (q = p; (q = strstr(q, "<sl:policy")); q++) {
            int type;
            if (!strncmp(q, "<sl:policyInt", 13)) type = SL_DATA_DWORD;
            else if (!strncmp(q, "<sl:policyStr", 13)) type = SL_DATA_SZ;
            else if (!strncmp(q, "<sl:policyBin", 13)) type = SL_DATA_BINARY;
            else continue;
            const char *nm = strstr(q, "name=\""), *gt = strchr(q, '>');
            if (!nm || !gt || nm > gt) continue;
            nm += 6;
            const char *nme = strchr(nm, '"');
            if (!nme || nme > gt) continue;
            const char *ve = strstr(gt + 1, "</sl:policy");
            if (!ve) continue;
            if (L->npol >= cap) {
                int ncap = cap ? cap * 2 : 64;
                struct policy *np = (struct policy *)LocalAlloc(LMEM_FIXED | LMEM_ZEROINIT, ncap * sizeof(*np));
                if (!np) break;
                if (L->pol) { memcpy(np, L->pol, L->npol * sizeof(*np)); LocalFree(L->pol); }
                L->pol = np; cap = ncap;
            }
            struct policy *P = &L->pol[L->npol];
            size_t nl = (size_t)(nme - nm), vl = (size_t)(ve - gt - 1);
            P->name = (char *)LocalAlloc(LMEM_FIXED, nl + 1);
            P->value = (char *)LocalAlloc(LMEM_FIXED, vl + 1);
            if (!P->name || !P->value) break;
            memcpy(P->name, nm, nl); P->name[nl] = 0;
            memcpy(P->value, gt + 1, vl); P->value[vl] = 0;
            P->type = type;
            L->npol++;
        }
    }
    for (q = p; (q = strstr(q, "<sl:appId")); q++) {
        const char *g = strstr(q, "<sl:guid>");
        const char *end = strstr(q, "</sl:appId>");
        if (g && end && g < end) { SLID id; if (parse_guid(g + 9, &id)) add_id(L->apps, &L->napps, &id); }
    }
    for (q = p; (q = strstr(q, "<sl:productIdRange")); q++) {
        const char *v = strstr(q, "value=\"");
        if (v && v - q < 120) { SLID id; if (parse_guid(v + 7, &id)) add_id(L->skus, &L->nskus, &id); }
    }
    if ((q = strstr(p, "msft:sl/grace-period"))) {
        /* <sl:floatingPeriod><sl:duration>P10D</sl:duration><sl:type>msft:sl/grace-period */
        const char *fp = q;
        while (fp > p && !(fp[0] == '<' && strncmp(fp, "<sl:floatingPeriod", 18) == 0)) fp--;
        const char *d = strstr(fp, "<sl:duration>P");
        if (d && d < q) L->grace_days = (DWORD)strtoul(d + 14, NULL, 10);
    }
    return 1;
}

static void init_store_dir(void)
{
    char base[MAX_PATH] = "";
    if (!GetEnvironmentVariableA("ProgramData", base, sizeof(base)) || !base[0]) lstrcpyA(base, "C:\\ProgramData");
    wsprintfA(store_dir, "%s\\ms365-spp", base);
    CreateDirectoryA(store_dir, NULL);
}

static char *read_file(const char *path, DWORD *size)
{
    HANDLE h = CreateFileA(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) return NULL;
    DWORD sz = GetFileSize(h, NULL), rd = 0;
    char *buf = (char *)LocalAlloc(LMEM_FIXED, sz + 1);
    if (!buf || !ReadFile(h, buf, sz, &rd, NULL)) { CloseHandle(h); if (buf) LocalFree(buf); return NULL; }
    CloseHandle(h);
    buf[rd] = 0;
    if (size) *size = rd;
    return buf;
}

static void load_store(void)
{
    if (nstore >= 0) return;
    nstore = 0;
    init_store_dir();
    char pattern[MAX_PATH];
    wsprintfA(pattern, "%s\\*.xrm-ms", store_dir);
    WIN32_FIND_DATAA fd;
    HANDLE h = FindFirstFileA(pattern, &fd);
    if (h == INVALID_HANDLE_VALUE) { tracef("sppc shim: store %s is empty", store_dir); return; }
    do {
        if (nstore >= MAX_LIC) break;
        char path[MAX_PATH];
        wsprintfA(path, "%s\\%s", store_dir, fd.cFileName);
        DWORD sz; char *xml = read_file(path, &sz);
        if (!xml) continue;
        if (parse_license(xml, sz, &store[nstore])) { lstrcpyA(store[nstore].path, path); nstore++; }
        LocalFree(xml);
    } while (FindNextFileA(h, &fd));
    FindClose(h);
    tracef("sppc shim: loaded %d licences from %s", nstore, store_dir);
}

/* ---- out-of-box grace period ---------------------------------------------
 * Office's OOB "Grace" licence (family *_Grace, floatingPeriod msft:sl/grace-period, 5 days) is what
 * lets a fresh, unactivated install run until it is activated. Windows starts that timer the first
 * time the right is consumed; we do the same and persist the start time next to the licences. */
static int is_grace_family(const struct lic *L) { return L->grace_days && strstr(L->family, "_Grace") != NULL; }

static ULONGLONG now_100ns(void) { FILETIME ft; GetSystemTimeAsFileTime(&ft); return ((ULONGLONG)ft.dwHighDateTime << 32) | ft.dwLowDateTime; }

/* remaining grace minutes for a SKU (starts the timer on first use); 0 = expired / no grace */
static DWORD grace_remaining_minutes(const SLID *sku, DWORD grace_days, int start_if_missing)
{
    char gs[48], path[MAX_PATH]; guid_str(sku, gs);
    wsprintfA(path, "%s\\grace-%s.txt", store_dir, gs);
    ULONGLONG start = 0; DWORD sz; char *txt = read_file(path, &sz);
    if (txt) { start = 0; for (char *c = txt; *c >= '0' && *c <= '9'; c++) start = start * 10 + (*c - '0'); LocalFree(txt); }
    ULONGLONG now = now_100ns();
    if (!start) {
        if (!start_if_missing) return grace_days * 24 * 60;
        start = now;
        char buf[32]; int n = 0; ULONGLONG v = start; char tmp[32]; int t = 0;
        do { tmp[t++] = (char)('0' + v % 10); v /= 10; } while (v);
        while (t) buf[n++] = tmp[--t];
        buf[n] = 0;
        HANDLE f = CreateFileA(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        DWORD wr; if (f != INVALID_HANDLE_VALUE) { WriteFile(f, buf, n, &wr, NULL); CloseHandle(f); }
        tracef("sppc shim: grace period started for %s (%lu days)", gs, grace_days);
    }
    ULONGLONG total = (ULONGLONG)grace_days * 24 * 60 * 60 * 10000000ULL;
    if (now < start) return grace_days * 24 * 60;
    if (now - start >= total) return 0;
    return (DWORD)((total - (now - start)) / (60 * 10000000ULL));
}

static struct lic *find_license(const SLID *id)
{
    for (int i = 0; i < nstore; i++) if (guid_eq(&store[i].id, id)) return &store[i];
    return NULL;
}
static int lic_has_app(const struct lic *L, const SLID *app)
{
    if (!app) return 1;
    for (int i = 0; i < L->napps; i++) if (guid_eq(&L->apps[i], app)) return 1;
    return 0;
}
static int lic_has_sku(const struct lic *L, const SLID *sku)
{
    if (!sku) return 1;
    for (int i = 0; i < L->nskus; i++) if (guid_eq(&L->skus[i], sku)) return 1;
    return 0;
}
/* the "ul-oob" licence for a SKU carries its family name and grace period */
static struct lic *sku_main_license(const SLID *sku)
{
    struct lic *best = NULL;
    for (int i = 0; i < nstore; i++)
        if (lic_has_sku(&store[i], sku) && store[i].nskus) { if (store[i].family[0]) return &store[i]; if (!best) best = &store[i]; }
    return best;
}

/* return a LocalAlloc'd wide string as SL_DATA_SZ */
static HRESULT ret_sz(const char *s, int *type, UINT *size, PBYTE *data)
{
    int n = MultiByteToWideChar(CP_UTF8, 0, s, -1, NULL, 0);
    WCHAR *w = (WCHAR *)LocalAlloc(LMEM_FIXED, n * sizeof(WCHAR));
    if (!w) return E_OUTOFMEMORY;
    MultiByteToWideChar(CP_UTF8, 0, s, -1, w, n);
    if (type) *type = SL_DATA_SZ;
    if (size) *size = n * sizeof(WCHAR);
    if (data) *data = (PBYTE)w; else LocalFree(w);
    return S_OK;
}
/* MULTI_SZ from an array of ASCII strings */
static HRESULT ret_multi_sz(const char **items, int n, int *type, UINT *size, PBYTE *data)
{
    UINT total = 1;
    for (int i = 0; i < n; i++) total += (UINT)strlen(items[i]) + 1;
    WCHAR *w = (WCHAR *)LocalAlloc(LMEM_FIXED | LMEM_ZEROINIT, total * sizeof(WCHAR));
    if (!w) return E_OUTOFMEMORY;
    WCHAR *p = w;
    for (int i = 0; i < n; i++) { int len = (int)strlen(items[i]); for (int k = 0; k <= len; k++) p[k] = (WCHAR)items[i][k]; p += len + 1; }
    *p = 0;
    if (type) *type = SL_DATA_MULTI_SZ;
    if (size) *size = total * sizeof(WCHAR);
    if (data) *data = (PBYTE)w; else LocalFree(w);
    return S_OK;
}

static HRESULT ret_dword(DWORD v, int *type, UINT *size, PBYTE *data)
{
    DWORD *d = (DWORD *)LocalAlloc(LMEM_FIXED, sizeof(DWORD));
    if (!d) return E_OUTOFMEMORY;
    *d = v;
    if (type) *type = SL_DATA_DWORD;
    if (size) *size = sizeof(DWORD);
    if (data) *data = (PBYTE)d; else LocalFree(d);
    return S_OK;
}
static int name_is(PCWSTR name, const char *ascii)
{
    if (!name) return 0;
    for (; *name && *ascii; name++, ascii++) if (*name != (WCHAR)*ascii) return 0;
    return !*name && !*ascii;
}
static int wname_to_ascii(PCWSTR w, char *out, size_t cap)
{
    size_t i = 0;
    if (!w) { out[0] = 0; return 0; }
    for (; w[i] && i + 1 < cap; i++) out[i] = (w[i] < 128) ? (char)w[i] : '?';
    out[i] = 0;
    return 1;
}

/* ---- policies (from the licences' <sl:policy*> elements) ---------------------
 * Policies that apply to a SKU come from every licence that names that SKU id (ul-oob) or its
 * family (PPD policy definitions, publishing licences). */
static const struct policy *find_policy(const SLID *sku, const char *name)
{
    const struct lic *main_l = NULL;
    for (int i = 0; i < nstore; i++)
        if (lic_has_sku(&store[i], sku) && store[i].nskus && store[i].family[0]) { main_l = &store[i]; break; }
    for (int i = 0; i < nstore; i++) {
        const struct lic *L = &store[i];
        int applies = (L->nskus && lic_has_sku(L, sku)) || (main_l && L->family[0] && !strcmp(L->family, main_l->family));
        if (!applies) continue;
        for (int k = 0; k < L->npol; k++) if (!strcmp(L->pol[k].name, name)) return &L->pol[k];
    }
    return NULL;
}
static DWORD parse_dword(const char *v)
{
    if (v[0] == '0' && (v[1] == 'x' || v[1] == 'X')) {
        DWORD r = 0;
        for (v += 2; *v; v++) { int h = hexval(*v); if (h < 0) break; r = r * 16 + h; }
        return r;
    }
    return (DWORD)strtoul(v, NULL, 10);
}
static HRESULT ret_bin_b64(const char *b64, int *type, UINT *size, PBYTE *data)
{
    static const char tbl[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    size_t n = strlen(b64);
    BYTE *out = (BYTE *)LocalAlloc(LMEM_FIXED, n * 3 / 4 + 4);
    if (!out) return E_OUTOFMEMORY;
    UINT o = 0; int val = 0, bits = 0;
    for (size_t i = 0; i < n; i++) {
        if (b64[i] == '=') break;
        const char *pp = strchr(tbl, b64[i]);
        if (!pp) continue;
        val = (val << 6) | (int)(pp - tbl); bits += 6;
        if (bits >= 8) { bits -= 8; out[o++] = (BYTE)((val >> bits) & 0xff); }
    }
    if (type) *type = SL_DATA_BINARY;
    if (size) *size = o;
    if (data) *data = out; else LocalFree(out);
    return S_OK;
}
static HRESULT ret_policy(const struct policy *P, int *type, UINT *size, PBYTE *data)
{
    if (P->type == SL_DATA_DWORD) return ret_dword(parse_dword(P->value), type, size, data);
    if (P->type == SL_DATA_BINARY) return ret_bin_b64(P->value, type, size, data);
    return ret_sz(P->value, type, size, data);
}
/* the SKU an application-policy handle refers to: the one given / consumed, else the grace SKU */
static SLID policy_sku; static int policy_sku_set;
static const SLID *current_sku(void)
{
    if (policy_sku_set) return &policy_sku;
    for (int i = 0; i < nstore; i++)
        if (is_grace_family(&store[i]) && store[i].nskus) { policy_sku = store[i].skus[0]; policy_sku_set = 1; return &policy_sku; }
    return NULL;
}
static HRESULT policy_lookup(PCWSTR name, int *type, UINT *size, PBYTE *data)
{
    if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL;
    char nm[200]; wname_to_ascii(name, nm, sizeof(nm));
    HRESULT hr = SL_E_VALUE_NOT_FOUND;
    EnterCriticalSection(&lock);
    load_store();
    const SLID *sku = current_sku();
    if (sku && !strcmp(nm, "*")) {
        /* enumerate: MULTI_SZ of every policy name that applies to the SKU */
        const char *names[2048]; int n = 0;
        for (int i = 0; i < nstore && n < 2048; i++)
            for (int k = 0; k < store[i].npol && n < 2048; k++)
                if (find_policy(sku, store[i].pol[k].name) == &store[i].pol[k]) names[n++] = store[i].pol[k].name;
        hr = ret_multi_sz(names, n, type, size, data);
    } else if (sku) {
        const struct policy *P = find_policy(sku, nm);
        if (P) hr = ret_policy(P, type, size, data);
    }
    LeaveCriticalSection(&lock);
    return hr;
}

/* ---- session ------------------------------------------------------------ */
API SLOpen(HSLC *handle)
{
    if (!handle) return E_INVALIDARG;
    EnterCriticalSection(&lock); load_store(); LeaveCriticalSection(&lock);
    *handle = (HSLC)(ULONG_PTR)0x534c4f50;
    return S_OK;
}
API SLClose(HSLC handle) { (void)handle; tracef("sppc shim: SLClose"); return S_OK; }

/* ---- licence installation ---------------------------------------------- */
API SLInstallLicense(HSLC h, UINT cb, const BYTE *blob, SLID *file_id)
{
    (void)h;
    if (!blob || !cb) return E_INVALIDARG;
    EnterCriticalSection(&lock);
    load_store();
    char *xml = (char *)LocalAlloc(LMEM_FIXED, cb + 1);
    if (!xml) { LeaveCriticalSection(&lock); return E_OUTOFMEMORY; }
    memcpy(xml, blob, cb); xml[cb] = 0;
    struct lic L;
    if (!parse_license(xml, cb, &L)) { LocalFree(xml); LeaveCriticalSection(&lock); tracef("sppc shim: SLInstallLicense: unparseable %u-byte blob", cb); return E_INVALIDARG; }
    wsprintfA(L.path, "%s\\%s.xrm-ms", store_dir, L.idstr);
    HANDLE f = CreateFileA(L.path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    DWORD wr = 0;
    if (f != INVALID_HANDLE_VALUE) { WriteFile(f, blob, cb, &wr, NULL); CloseHandle(f); }
    LocalFree(xml);
    struct lic *slot = find_license(&L.id);
    if (!slot && nstore < MAX_LIC) slot = &store[nstore++];
    if (slot) *slot = L;
    if (file_id) *file_id = L.id;
    tracef("sppc shim: SLInstallLicense %u bytes -> %s \"%s\" family=%s skus=%d apps=%d grace=%lu",
           cb, L.idstr, L.title, L.family, L.nskus, L.napps, L.grace_days);
    LeaveCriticalSection(&lock);
    return S_OK;
}
API SLUninstallLicense(HSLC h, const SLID *id)
{
    (void)h;
    EnterCriticalSection(&lock);
    load_store();
    struct lic *L = find_license(id);
    if (L) { DeleteFileA(L->path); *L = store[--nstore]; }
    LeaveCriticalSection(&lock);
    TRACE_RET("SLUninstallLicense", id, NULL, NULL, S_OK);
}
API SLGetLicense(HSLC h, const SLID *file_id, UINT *size, PBYTE *data)
{
    (void)h;
    EnterCriticalSection(&lock);
    load_store();
    struct lic *L = find_license(file_id);
    HRESULT hr = SL_E_VALUE_NOT_FOUND;
    if (L) {
        DWORD sz; char *blob = read_file(L->path, &sz);
        if (blob) { if (size) *size = sz; if (data) *data = (PBYTE)blob; else LocalFree(blob); hr = S_OK; }
    }
    LeaveCriticalSection(&lock);
    TRACE_RET("SLGetLicense", file_id, NULL, NULL, hr);
}
API SLGetLicenseFileId(HSLC h, UINT cb, const BYTE *blob, SLID *file_id)
{
    (void)h;
    char *xml = (char *)LocalAlloc(LMEM_FIXED, cb + 1);
    if (!xml) return E_OUTOFMEMORY;
    memcpy(xml, blob, cb); xml[cb] = 0;
    struct lic L; int ok = parse_license(xml, cb, &L);
    LocalFree(xml);
    if (ok && file_id) *file_id = L.id;
    return ok ? S_OK : SL_E_VALUE_NOT_FOUND;
}

/* ---- enumeration --------------------------------------------------------- */
API SLGetSLIDList(HSLC h, int qtype, const SLID *qid, int rtype, UINT *count, SLID **ids)
{
    (void)h;
    SLID out[MAX_LIC]; int n = 0;
    EnterCriticalSection(&lock);
    load_store();
    for (int i = 0; i < nstore && n < MAX_LIC; i++) {
        const struct lic *L = &store[i];
        int match = 0;
        switch (qtype) {
        case SL_ID_APPLICATION: match = lic_has_app(L, qid); break;
        case SL_ID_PRODUCT_SKU: match = lic_has_sku(L, qid); break;
        case SL_ID_LICENSE: case SL_ID_LICENSE_FILE: match = guid_eq(&L->id, qid); break;
        case SL_ID_ALL_LICENSES: case SL_ID_ALL_LICENSE_FILES: match = 1; break;
        default: match = 0;
        }
        if (!match) continue;
        switch (rtype) {
        case SL_ID_PRODUCT_SKU: for (int k = 0; k < L->nskus; k++) add_id(out, &n, &L->skus[k]); break;
        case SL_ID_APPLICATION: for (int k = 0; k < L->napps; k++) add_id(out, &n, &L->apps[k]); break;
        case SL_ID_LICENSE: case SL_ID_LICENSE_FILE: case SL_ID_ALL_LICENSES: case SL_ID_ALL_LICENSE_FILES: add_id(out, &n, &L->id); break;
        default: break; /* SL_ID_PKEY, store tokens: none */
        }
    }
    LeaveCriticalSection(&lock);
    SLID *arr = NULL;
    if (n) { arr = (SLID *)LocalAlloc(LMEM_FIXED, n * sizeof(SLID)); if (!arr) return E_OUTOFMEMORY; memcpy(arr, out, n * sizeof(SLID)); }
    if (count) *count = n;
    if (ids) *ids = arr; else if (arr) LocalFree(arr);
    tracef("sppc shim: SLGetSLIDList qtype=%d rtype=%d -> %d ids", qtype, rtype, n);
    return S_OK;
}
API SLGetInstalledProductKeyIds(HSLC h, const SLID *product, UINT *count, SLID **ids)
{
    (void)h;
    if (count) *count = 0;
    if (ids) *ids = NULL;
    TRACE_RET("SLGetInstalledProductKeyIds", product, NULL, NULL, S_OK);
}

/* ---- information getters ------------------------------------------------- */
API SLGetProductSkuInformation(HSLC h, const SLID *sku, PCWSTR name, int *type, UINT *size, PBYTE *data)
{
    (void)h;
    HRESULT hr = SL_E_VALUE_NOT_FOUND;
    if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL;
    EnterCriticalSection(&lock);
    load_store();
    struct lic *L = sku_main_license(sku);
    if (L) {
        char nm[200], pref[220];
        wname_to_ascii(name, nm, sizeof(nm));
        wsprintfA(pref, "office-%s", nm);
        const struct policy *P = find_policy(sku, pref);
        if (!P) P = find_policy(sku, nm);
        if (name_is(name, "Name"))        hr = ret_sz(L->family[0] ? L->family : L->title, type, size, data);
        else if (name_is(name, "Description")) hr = ret_sz(L->title, type, size, data);
        else if (name_is(name, "Family") || name_is(name, "Security-SPP-Reserved-Family") || name_is(name, "LicenseFamily")) hr = ret_sz(L->family, type, size, data);
        else if (P) hr = ret_policy(P, type, size, data);
    }
    LeaveCriticalSection(&lock);
    TRACE_RET("SLGetProductSkuInformation", sku, NULL, name, hr);
}
API SLGetLicenseInformation(HSLC h, const SLID *id, PCWSTR name, int *type, UINT *size, PBYTE *data)
{
    (void)h;
    HRESULT hr = SL_E_VALUE_NOT_FOUND;
    if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL;
    EnterCriticalSection(&lock);
    load_store();
    struct lic *L = find_license(id);
    if (L) {
        char nm[200]; wname_to_ascii(name, nm, sizeof(nm));
        if (name_is(name, "Name") || name_is(name, "Description")) hr = ret_sz(L->title, type, size, data);
        else if (name_is(name, "Family")) hr = ret_sz(L->family, type, size, data);
        else for (int k = 0; k < L->npol && hr != S_OK; k++)
            if (!strcmp(L->pol[k].name, nm) || (!strncmp(L->pol[k].name, "office-", 7) && !strcmp(L->pol[k].name + 7, nm)))
                hr = ret_policy(&L->pol[k], type, size, data);
    }
    LeaveCriticalSection(&lock);
    TRACE_RET("SLGetLicenseInformation", id, NULL, name, hr);
}
API SLGetApplicationInformation(HSLC h, const SLID *app, PCWSTR name, int *type, UINT *size, PBYTE *data)
{
    (void)h;
    if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL;
    TRACE_RET("SLGetApplicationInformation", app, NULL, name, SL_E_VALUE_NOT_FOUND);
}
API SLGetPKeyInformation(HSLC h, const SLID *pkey, PCWSTR name, int *type, UINT *size, PBYTE *data)
{
    (void)h;
    if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL;
    TRACE_RET("SLGetPKeyInformation", pkey, NULL, name, SL_E_VALUE_NOT_FOUND);
}
API SLGetServiceInformation(HSLC h, PCWSTR name, int *type, UINT *size, PBYTE *data)
{
    (void)h;
    HRESULT hr = SL_E_VALUE_NOT_FOUND;
    if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL;
    if (name_is(name, "Version")) hr = ret_sz("10.0.19041.1", type, size, data);
    else if (name_is(name, "IsKeyManagementServiceMachine") || name_is(name, "KeyManagementServiceListeningPort")) hr = ret_dword(0, type, size, data);
    else if (name_is(name, "ActivePlugins")) {
        /* Office's SPP plug-ins: the Click-to-Run licensing and extensibility components (mso30win32client
         * keeps these GUIDs right next to the "ActivePlugins" query and fails with SL_E_PLUGIN_NOT_FOUND otherwise) */
        static const char *plugins[] = { "{90160000-007E-0000-1000-0000000FF1CE}", "{90160000-008F-0000-1000-0000000FF1CE}" };
        hr = ret_multi_sz(plugins, 2, type, size, data);
    }
    TRACE_RET("SLGetServiceInformation", NULL, NULL, name, hr);
}
API SLGetPolicyInformation(HSLP h, PCWSTR name, int *type, UINT *size, PBYTE *data)
{ (void)h; TRACE_RET("SLGetPolicyInformation", NULL, NULL, name, policy_lookup(name, type, size, data)); }
API SLGetApplicationPolicy(HSLP h, PCWSTR name, int *type, UINT *size, PBYTE *data)
{ (void)h; TRACE_RET("SLGetApplicationPolicy", NULL, NULL, name, policy_lookup(name, type, size, data)); }
API SLGetPolicyInformationDWORD(HSLP h, PCWSTR name, DWORD *out)
{
    (void)h;
    int type = 0; UINT size = 0; PBYTE data = NULL;
    HRESULT hr = policy_lookup(name, &type, &size, &data);
    if (hr == S_OK && type == SL_DATA_DWORD && size == sizeof(DWORD)) { if (out) *out = *(DWORD *)data; }
    else if (hr == S_OK) hr = 0xC004F01E; /* SL_E_DATATYPE_MISMATCHED */
    if (data) LocalFree(data);
    if (hr != S_OK && out) *out = 0;
    TRACE_RET("SLGetPolicyInformationDWORD", NULL, NULL, name, hr);
}
API SLGetGenuineInformation(const SLID *id, PCWSTR name, int *type, UINT *size, PBYTE *data)
{ if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL; TRACE_RET("SLGetGenuineInformation", id, NULL, name, SL_E_VALUE_NOT_FOUND); }

/* ---- status / rights: installed but not activated ------------------------ */
API SLGetLicensingStatusInformation(HSLC h, const SLID *app, const SLID *sku, PCWSTR right, UINT *count, SL_LICENSING_STATUS **status)
{
    (void)h; (void)right;
    SLID skus[MAX_LIC]; int n = 0;
    EnterCriticalSection(&lock);
    load_store();
    for (int i = 0; i < nstore; i++) {
        const struct lic *L = &store[i];
        if (!lic_has_app(L, app) || !lic_has_sku(L, sku)) continue;
        for (int k = 0; k < L->nskus; k++) if (!sku || guid_eq(&L->skus[k], sku)) add_id(skus, &n, &L->skus[k]);
    }
    SL_LICENSING_STATUS *arr = NULL;
    if (n) {
        arr = (SL_LICENSING_STATUS *)LocalAlloc(LMEM_FIXED | LMEM_ZEROINIT, n * sizeof(*arr));
        if (arr) for (int i = 0; i < n; i++) {
            struct lic *L = sku_main_license(&skus[i]);
            arr[i].SkuId = skus[i];
            arr[i].dwTotalGraceDays = L ? L->grace_days : 0;
            arr[i].qwValidityExpiration = 0;
            DWORD rem = (L && is_grace_family(L)) ? grace_remaining_minutes(&skus[i], L->grace_days, 0) : 0;
            /* status reason for a grace-period product is SL_E_GRACE_PERIOD-style "running within the valid grace period" */
            if (rem) { arr[i].eStatus = SL_LICENSING_STATUS_IN_GRACE_PERIOD; arr[i].dwGraceTime = rem; arr[i].hrReason = (HRESULT)0xC004F00C; }
            else if (L && is_grace_family(L)) { arr[i].eStatus = SL_LICENSING_STATUS_NOTIFICATION; arr[i].dwGraceTime = 0; arr[i].hrReason = 0xC004F009; /* grace expired */ }
            else { arr[i].eStatus = SL_LICENSING_STATUS_UNLICENSED; arr[i].dwGraceTime = 0; arr[i].hrReason = SL_E_PKEY_NOT_INSTALLED; }
        }
    }
    LeaveCriticalSection(&lock);
    if (n && !arr) return E_OUTOFMEMORY;
    if (count) *count = n;
    if (status) *status = arr; else if (arr) LocalFree(arr);
    char ga[48], gb[48]; guid_str(app, ga); guid_str(sku, gb);
    tracef("sppc shim: SLGetLicensingStatusInformation app=%s sku=%s -> %d statuses", ga, gb, n);
    return S_OK;
}
API SLConsumeRight(HSLC h, const SLID *app, const SLID *sku, PCWSTR right, void *reserved)
{
    (void)h; (void)reserved;
    HRESULT hr = SL_E_RIGHT_NOT_GRANTED;
    EnterCriticalSection(&lock);
    load_store();
    for (int i = 0; i < nstore && hr != S_OK; i++) {
        const struct lic *L = &store[i];
        if (!is_grace_family(L) || !lic_has_app(L, app) || !lic_has_sku(L, sku)) continue;
        for (int k = 0; k < L->nskus; k++)
            if ((!sku || guid_eq(&L->skus[k], sku)) && grace_remaining_minutes(&L->skus[k], L->grace_days, 1)) { hr = S_OK; policy_sku = L->skus[k]; policy_sku_set = 1; break; }
    }
    LeaveCriticalSection(&lock);
    TRACE_RET("SLConsumeRight", app, sku, right, hr);
}
API SLGetActiveLicenseInfo(HSLC h, const SLID *app, const SLID *sku, UINT *count, void **info)
{ (void)h; if (count) *count = 0; if (info) *info = NULL; TRACE_RET("SLGetActiveLicenseInfo", app, sku, NULL, SL_E_VALUE_NOT_FOUND); }

/* ---- product keys: none can be installed here ----------------------------- */
API SLInstallProofOfPurchase(HSLC h, PCWSTR alg, PCWSTR key, UINT cb, PBYTE data, SLID *pkey_id)
{ (void)h; (void)alg; (void)cb; (void)data; if (pkey_id) memset(pkey_id, 0, sizeof(*pkey_id)); tracef("sppc shim: SLInstallProofOfPurchase %ls (refused)", key ? key : L""); return SL_E_VALUE_NOT_FOUND; }
API SLInstallProofOfPurchaseEx(HSLC h, const SLID *app, PCWSTR alg, PCWSTR key, UINT cb, PBYTE data, SLID *pkey_id)
{ (void)h; (void)app; (void)alg; (void)cb; (void)data; if (pkey_id) memset(pkey_id, 0, sizeof(*pkey_id)); tracef("sppc shim: SLInstallProofOfPurchaseEx %ls (refused)", key ? key : L""); return SL_E_VALUE_NOT_FOUND; }
API SLUninstallProofOfPurchase(HSLC h, const SLID *id) { (void)h; TRACE_RET("SLUninstallProofOfPurchase", id, NULL, NULL, S_OK); }
API SLSetCurrentProductKey(HSLC h, const SLID *sku, const SLID *pkey) { (void)h; TRACE_RET("SLSetCurrentProductKey", sku, pkey, NULL, SL_E_VALUE_NOT_FOUND); }
API SLGetPKeyId(HSLC h, PCWSTR alg, PCWSTR key, UINT cb, const BYTE *data, SLID *pkey_id)
{ (void)h; (void)alg; (void)cb; (void)data; if (pkey_id) memset(pkey_id, 0, sizeof(*pkey_id)); TRACE_RET("SLGetPKeyId", NULL, NULL, key, SL_E_VALUE_NOT_FOUND); }
API SLGetEncryptedPIDEx(HSLC h, const SLID *sku, UINT *size, PBYTE *data)
{ (void)h; if (size) *size = 0; if (data) *data = NULL; TRACE_RET("SLGetEncryptedPIDEx", sku, NULL, NULL, SL_E_VALUE_NOT_FOUND); }
/* Authentication data: Office hands SPP a blob before consuming rights; the SPP plug-in evaluates the
 * licence's external validator against it and Office reads the outcome back with
 * SLGetAuthenticationResult. Kept per process; dumped to the trace so the format can be studied. */
static BYTE auth_data[1024]; static UINT auth_len;
API SLSetAuthenticationData(HSLC h, UINT cb, const BYTE *pb)
{
    (void)h;
    if (pb && cb && cb <= sizeof(auth_data)) { memcpy(auth_data, pb, cb); auth_len = cb; }
    char hex[3 * 64 + 8]; UINT n = cb < 64 ? cb : 64; hex[0] = 0;
    for (UINT i = 0; pb && i < n; i++) wsprintfA(hex + 3 * i, "%02x ", pb[i]);
    tracef("sppc shim: SLSetAuthenticationData %u bytes: %s", cb, hex);
    if (pb && cb > 64) {
        for (UINT off = 64; off < cb; off += 64) {
            UINT m = (cb - off) < 64 ? (cb - off) : 64; hex[0] = 0;
            for (UINT i = 0; i < m; i++) wsprintfA(hex + 3 * i, "%02x ", pb[off + i]);
            tracef("sppc shim:   +%03x: %s", off, hex);
        }
    }
    return S_OK;
}
API SLGetAuthenticationResult(HSLC h, UINT *size, PBYTE *data)
{
    (void)h;
    HRESULT hr = SL_E_VALUE_NOT_FOUND;
    if (size) *size = 0; if (data) *data = NULL;
    if (auth_len) {
        /* experiment: hand the authentication data back as the result */
        BYTE *out = (BYTE *)LocalAlloc(LMEM_FIXED, auth_len);
        if (out) { memcpy(out, auth_data, auth_len); if (size) *size = auth_len; if (data) *data = out; else LocalFree(out); hr = S_OK; }
    }
    TRACE_RET("SLGetAuthenticationResult", NULL, NULL, NULL, hr);
}
API SLGenerateOfflineInstallationId(HSLC h, const SLID *sku, PWSTR *out)
{ (void)h; if (out) *out = NULL; TRACE_RET("SLGenerateOfflineInstallationId", sku, NULL, NULL, SL_E_VALUE_NOT_FOUND); }
API SLGenerateOfflineInstallationIdEx(HSLC h, const SLID *sku, const void *info, PWSTR *out)
{ (void)h; (void)info; if (out) *out = NULL; TRACE_RET("SLGenerateOfflineInstallationIdEx", sku, NULL, NULL, SL_E_VALUE_NOT_FOUND); }
API SLGatherMigrationBlob(BOOL a, BOOL b, UINT *size, PBYTE data) { (void)a; (void)b; (void)data; if (size) *size = 0; TRACE_RET("SLGatherMigrationBlob", NULL, NULL, NULL, S_OK); }
API SLGatherMigrationBlobEx(BOOL a, BOOL b, UINT *size, PBYTE data) { (void)a; (void)b; (void)data; if (size) *size = 0; TRACE_RET("SLGatherMigrationBlobEx", NULL, NULL, NULL, S_OK); }
API SLIsGenuineLocalEx(const SLID *app, const SLID *alt, int *state) { if (state) *state = 0; TRACE_RET("SLIsGenuineLocalEx", app, alt, NULL, S_OK); }

/* ---- policies / events: accept ------------------------------------------- */
API SLLoadApplicationPolicies(const SLID *app, const SLID *sku, DWORD flags, HSLP *handle)
{
    (void)flags;
    EnterCriticalSection(&lock);
    load_store();
    if (sku) { policy_sku = *sku; policy_sku_set = 1; }
    LeaveCriticalSection(&lock);
    if (handle) *handle = (HSLP)(ULONG_PTR)0x534c504f;
    TRACE_RET("SLLoadApplicationPolicies", app, sku, NULL, S_OK);
}
API SLUnloadApplicationPolicies(HSLP handle, DWORD flags) { (void)handle; (void)flags; TRACE_RET("SLUnloadApplicationPolicies", NULL, NULL, NULL, S_OK); }
API SLPersistApplicationPolicies(const SLID *app, const SLID *sku, DWORD flags) { (void)flags; TRACE_RET("SLPersistApplicationPolicies", app, sku, NULL, S_OK); }
API SLRegisterEvent(HSLC h, PCWSTR name, const SLID *app, HANDLE event) { (void)h; (void)event; TRACE_RET("SLRegisterEvent", app, NULL, name, S_OK); }
API SLUnregisterEvent(HSLC h, PCWSTR name, const SLID *app, HANDLE event) { (void)h; (void)event; TRACE_RET("SLUnregisterEvent", app, NULL, name, S_OK); }

#define OK_STUB(name) API name(void *a, void *b, void *c, void *d) { tracef("sppc shim: " #name "(%p, %p, %p, %p) -> S_OK", a, b, c, d); return S_OK; }
OK_STUB(SLCallServer)
OK_STUB(SLDepositMigrationBlob)
OK_STUB(SLDepositOfflineConfirmationId)
OK_STUB(SLDepositOfflineConfirmationIdEx)
OK_STUB(SLDepositStoreToken)
OK_STUB(SLFireEvent)
OK_STUB(SLPersistRTSPayloadOverride)
OK_STUB(SLReArm)
API SLRegisterPlugin(void *a, void *b, void *c, void *d)
{
    tracef("sppc shim: SLRegisterPlugin(%p, %p, %p, %p)", a, b, c, d);
    void *args[4] = { a, b, c, d };
    for (int i = 0; i < 4; i++) {
        if (!args[i] || IsBadReadPtr(args[i], 2)) continue;
        const WCHAR *w = (const WCHAR *)args[i];
        if (!IsBadStringPtrW(w, 200) && w[0] > 31 && w[0] < 127 && w[1] > 31 && w[1] < 127) tracef("sppc shim:   arg%d = L\"%.100ls\"", i, w);
        else if (!IsBadReadPtr(args[i], 16)) { char g[48]; guid_str((const SLID *)args[i], g); tracef("sppc shim:   arg%d as guid = %s", i, g); }
    }
    return S_OK;
}
OK_STUB(SLSetGenuineInformation)
OK_STUB(SLUnregisterPlugin)
OK_STUB(SLpAuthenticateGenuineTicketResponse)
OK_STUB(SLpBeginGenuineTicketTransaction)
OK_STUB(SLpClearActivationInProgress)
OK_STUB(SLpDepositDownlevelGenuineTicket)
OK_STUB(SLpDepositTokenActivationResponse)
OK_STUB(SLpGenerateTokenActivationChallenge)
OK_STUB(SLpGetGenuineBlob)
OK_STUB(SLpGetGenuineLocal)
OK_STUB(SLpGetLicenseAcquisitionInfo)
OK_STUB(SLpGetMSPidInformation)
OK_STUB(SLpGetMachineUGUID)
OK_STUB(SLpGetTokenActivationGrantInfo)
OK_STUB(SLpIAActivateProduct)
OK_STUB(SLpIsCurrentInstalledProductKeyDefaultKey)
OK_STUB(SLpProcessVMPipeMessage)
OK_STUB(SLpSetActivationInProgress)
OK_STUB(SLpTriggerServiceWorker)
OK_STUB(SLpVLActivateProduct)

BOOL WINAPI DllMain(HINSTANCE inst, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) { DisableThreadLibraryCalls(inst); InitializeCriticalSection(&lock); }
    return TRUE;
}
