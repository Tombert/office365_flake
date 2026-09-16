/*
 * ms365uia.dll: in-process COM server for CLSID_CUIAutomationRegistrar
 * ({6e29fabf-9977-42d1-8d0e-ca7e61ad87e6}), which Wine's uiautomationcore lacks.
 *
 * Office creates this object when a document gets focus to register custom UI Automation
 * properties, events and patterns, and dereferences the result without checking for failure.
 * IUIAutomationRegistrar only has to hand back identifiers; nothing consumes them under Wine.
 */
#include <windows.h>
#include <objbase.h>
#include <string.h>

typedef int PROPERTYID, EVENTID, PATTERNID;

/* {8609c4ec-4a1a-4d88-a357-5a66e060e1cf} IUIAutomationRegistrar */
static const GUID IID_IUIAutomationRegistrar_ = { 0x8609c4ec, 0x4a1a, 0x4d88, { 0xa3, 0x57, 0x5a, 0x66, 0xe0, 0x60, 0xe1, 0xcf } };
/* {6e29fabf-9977-42d1-8d0e-ca7e61ad87e6} CLSID_CUIAutomationRegistrar */
static const GUID CLSID_Registrar_ = { 0x6e29fabf, 0x9977, 0x42d1, { 0x8d, 0x0e, 0xca, 0x7e, 0x61, 0xad, 0x87, 0xe6 } };

static LONG obj_count;
static LONG next_id = 60000;   /* custom IDs start above the well-known ranges */

/* ---- IUIAutomationRegistrar --------------------------------------------------------------- */
typedef struct registrar {
    const struct registrar_vtbl *lpVtbl;
    LONG ref;
} registrar;

struct registrar_vtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(registrar *, REFIID, void **);
    ULONG   (STDMETHODCALLTYPE *AddRef)(registrar *);
    ULONG   (STDMETHODCALLTYPE *Release)(registrar *);
    HRESULT (STDMETHODCALLTYPE *RegisterProperty)(registrar *, const void *property, PROPERTYID *id);
    HRESULT (STDMETHODCALLTYPE *RegisterEvent)(registrar *, const void *event, EVENTID *id);
    HRESULT (STDMETHODCALLTYPE *RegisterPattern)(registrar *, const void *pattern, PATTERNID *patternId,
                                                 PROPERTYID *pairedPatternAvailablePropertyId,
                                                 UINT propertyIdCount, PROPERTYID *propertyIds,
                                                 UINT eventIdCount, EVENTID *eventIds);
};

static HRESULT STDMETHODCALLTYPE reg_QueryInterface(registrar *This, REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    if (IsEqualGUID(riid, &IID_IUnknown) || IsEqualGUID(riid, &IID_IUIAutomationRegistrar_)) {
        *ppv = This; InterlockedIncrement(&This->ref); return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE reg_AddRef(registrar *This) { return InterlockedIncrement(&This->ref); }
static ULONG STDMETHODCALLTYPE reg_Release(registrar *This)
{
    LONG r = InterlockedDecrement(&This->ref);
    if (!r) { HeapFree(GetProcessHeap(), 0, This); InterlockedDecrement(&obj_count); }
    return r;
}
static HRESULT STDMETHODCALLTYPE reg_RegisterProperty(registrar *This, const void *property, PROPERTYID *id)
{
    (void)This; (void)property;
    if (!id) return E_POINTER;
    *id = InterlockedIncrement(&next_id);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE reg_RegisterEvent(registrar *This, const void *event, EVENTID *id)
{
    (void)This; (void)event;
    if (!id) return E_POINTER;
    *id = InterlockedIncrement(&next_id);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE reg_RegisterPattern(registrar *This, const void *pattern, PATTERNID *patternId,
                                                     PROPERTYID *pairedId, UINT nprops, PROPERTYID *props,
                                                     UINT nevents, EVENTID *events)
{
    (void)This; (void)pattern;
    if (patternId) *patternId = InterlockedIncrement(&next_id);
    if (pairedId) *pairedId = InterlockedIncrement(&next_id);
    for (UINT i = 0; props && i < nprops; i++) props[i] = InterlockedIncrement(&next_id);
    for (UINT i = 0; events && i < nevents; i++) events[i] = InterlockedIncrement(&next_id);
    return S_OK;
}
static const struct registrar_vtbl registrar_vtable = {
    reg_QueryInterface, reg_AddRef, reg_Release, reg_RegisterProperty, reg_RegisterEvent, reg_RegisterPattern
};

/* ---- IClassFactory ------------------------------------------------------------------------ */
typedef struct factory { const struct factory_vtbl *lpVtbl; } factory;
struct factory_vtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(factory *, REFIID, void **);
    ULONG   (STDMETHODCALLTYPE *AddRef)(factory *);
    ULONG   (STDMETHODCALLTYPE *Release)(factory *);
    HRESULT (STDMETHODCALLTYPE *CreateInstance)(factory *, IUnknown *outer, REFIID riid, void **ppv);
    HRESULT (STDMETHODCALLTYPE *LockServer)(factory *, BOOL lock);
};
static HRESULT STDMETHODCALLTYPE fac_QueryInterface(factory *This, REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    if (IsEqualGUID(riid, &IID_IUnknown) || IsEqualGUID(riid, &IID_IClassFactory)) { *ppv = This; return S_OK; }
    *ppv = NULL;
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE fac_AddRef(factory *This) { (void)This; return 2; }
static ULONG STDMETHODCALLTYPE fac_Release(factory *This) { (void)This; return 1; }
static HRESULT STDMETHODCALLTYPE fac_CreateInstance(factory *This, IUnknown *outer, REFIID riid, void **ppv)
{
    (void)This;
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (outer) return CLASS_E_NOAGGREGATION;
    registrar *r = (registrar *)HeapAlloc(GetProcessHeap(), 0, sizeof(*r));
    if (!r) return E_OUTOFMEMORY;
    r->lpVtbl = &registrar_vtable;
    r->ref = 1;
    InterlockedIncrement(&obj_count);
    HRESULT hr = reg_QueryInterface(r, riid, ppv);
    reg_Release(r);
    OutputDebugStringA("ms365uia: created IUIAutomationRegistrar");
    return hr;
}
static HRESULT STDMETHODCALLTYPE fac_LockServer(factory *This, BOOL lock) { (void)This; (void)lock; return S_OK; }
static const struct factory_vtbl factory_vtable = { fac_QueryInterface, fac_AddRef, fac_Release, fac_CreateInstance, fac_LockServer };
static factory the_factory = { &factory_vtable };

__declspec(dllexport) HRESULT WINAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (!IsEqualGUID(rclsid, &CLSID_Registrar_)) return CLASS_E_CLASSNOTAVAILABLE;
    return fac_QueryInterface(&the_factory, riid, ppv);
}
__declspec(dllexport) HRESULT WINAPI DllCanUnloadNow(void) { return obj_count ? S_FALSE : S_OK; }

BOOL WINAPI DllMain(HINSTANCE inst, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(inst);
    return TRUE;
}
