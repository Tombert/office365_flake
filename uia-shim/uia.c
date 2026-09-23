/*
 * ms365uia.dll: in-process COM server for CLSID_CUIAutomationRegistrar
 * ({6e29fabf-9977-42d1-8d0e-ca7e61ad87e6}), which Wine's uiautomationcore lacks.
 *
 * Office creates this object when a document gets focus to register custom UI Automation
 * properties, events and patterns, and dereferences the result without checking for failure.
 * IUIAutomationRegistrar only has to hand back identifiers; nothing consumes them under Wine.
 *
 * Also CLSID_InkDisp ({937c1a34-151d-4610-9ca6-a8cc9bdb5d83}), the Tablet PC ink object from
 * InkObj.dll, which Wine only has as a stub DLL without class objects. OneNote creates one at
 * start; when that fails it assumes a Windows Server without the "Desktop Experience" feature and
 * refuses to open. The object here is an empty ink document: it can be created, answers IUnknown,
 * IDispatch and IInkDisp, and every ink method reports E_NOTIMPL (no pen input under Wine anyway).
 */
#include <windows.h>
#include <objbase.h>
#include <string.h>

typedef int PROPERTYID, EVENTID, PATTERNID;

/* {8609c4ec-4a1a-4d88-a357-5a66e060e1cf} IUIAutomationRegistrar */
static const GUID IID_IUIAutomationRegistrar_ = { 0x8609c4ec, 0x4a1a, 0x4d88, { 0xa3, 0x57, 0x5a, 0x66, 0xe0, 0x60, 0xe1, 0xcf } };
/* {6e29fabf-9977-42d1-8d0e-ca7e61ad87e6} CLSID_CUIAutomationRegistrar */
static const GUID CLSID_Registrar_ = { 0x6e29fabf, 0x9977, 0x42d1, { 0x8d, 0x0e, 0xca, 0x7e, 0x61, 0xad, 0x87, 0xe6 } };

/* {937c1a34-151d-4610-9ca6-a8cc9bdb5d83} CLSID_InkDisp, {9d398fa0-c4e2-4fcd-9973-975caaf47ea6} IInkDisp */
static const GUID CLSID_InkDisp_ = { 0x937c1a34, 0x151d, 0x4610, { 0x9c, 0xa6, 0xa8, 0xcc, 0x9b, 0xdb, 0x5d, 0x83 } };
static const GUID IID_IInkDisp_  = { 0x9d398fa0, 0xc4e2, 0x4fcd, { 0x99, 0x73, 0x97, 0x5c, 0xaa, 0xf4, 0x7e, 0xa6 } };

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

/* ---- InkDisp ------------------------------------------------------------------------------ */
typedef struct ink { void *const *lpVtbl; LONG ref; } ink;
static HRESULT STDMETHODCALLTYPE ink_QueryInterface(ink *This, REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    if (IsEqualGUID(riid, &IID_IUnknown) || IsEqualGUID(riid, &IID_IDispatch) || IsEqualGUID(riid, &IID_IInkDisp_)) {
        *ppv = This; InterlockedIncrement(&This->ref); return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE ink_AddRef(ink *This) { return InterlockedIncrement(&This->ref); }
static ULONG STDMETHODCALLTYPE ink_Release(ink *This)
{
    LONG r = InterlockedDecrement(&This->ref);
    if (!r) { HeapFree(GetProcessHeap(), 0, This); InterlockedDecrement(&obj_count); }
    return r;
}
static HRESULT STDMETHODCALLTYPE ink_GetTypeInfoCount(ink *This, UINT *n) { (void)This; if (!n) return E_POINTER; *n = 0; return S_OK; }
static HRESULT STDMETHODCALLTYPE ink_GetIDsOfNames(void) { return DISP_E_UNKNOWNNAME; }
static HRESULT STDMETHODCALLTYPE ink_Invoke(void) { return DISP_E_MEMBERNOTFOUND; }
static HRESULT STDMETHODCALLTYPE ink_notimpl(void) { return E_NOTIMPL; }
/* IUnknown, IDispatch, then IInkDisp's properties and methods (fewer than 60); the Win64 ABI is
 * caller-clean, so one argument-ignoring stub serves every slot */
#define INK_SLOTS 7 + 60
static void *ink_vtable[INK_SLOTS];
static void ink_init_vtable(void)
{
    if (ink_vtable[0]) return;
    for (int i = 7; i < INK_SLOTS; i++) ink_vtable[i] = (void *)ink_notimpl;
    ink_vtable[3] = (void *)ink_GetTypeInfoCount;
    ink_vtable[4] = (void *)ink_notimpl;          /* GetTypeInfo */
    ink_vtable[5] = (void *)ink_GetIDsOfNames;
    ink_vtable[6] = (void *)ink_Invoke;
    ink_vtable[1] = (void *)ink_AddRef;
    ink_vtable[2] = (void *)ink_Release;
    ink_vtable[0] = (void *)ink_QueryInterface;   /* last: marks the table complete */
}
static HRESULT ink_create(REFIID riid, void **ppv)
{
    ink_init_vtable();
    ink *k = (ink *)HeapAlloc(GetProcessHeap(), 0, sizeof(*k));
    if (!k) return E_OUTOFMEMORY;
    k->lpVtbl = (void *const *)ink_vtable;
    k->ref = 1;
    InterlockedIncrement(&obj_count);
    HRESULT hr = ink_QueryInterface(k, riid, ppv);
    ink_Release(k);
    OutputDebugStringA("ms365uia: created InkDisp");
    return hr;
}

/* ---- IClassFactory ------------------------------------------------------------------------ */
typedef struct factory { const struct factory_vtbl *lpVtbl; int ink; } factory;
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
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (outer) return CLASS_E_NOAGGREGATION;
    if (This->ink) return ink_create(riid, ppv);
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
static factory registrar_factory = { &factory_vtable, 0 };
static factory ink_factory = { &factory_vtable, 1 };

__declspec(dllexport) HRESULT WINAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualGUID(rclsid, &CLSID_Registrar_)) return fac_QueryInterface(&registrar_factory, riid, ppv);
    if (IsEqualGUID(rclsid, &CLSID_InkDisp_)) return fac_QueryInterface(&ink_factory, riid, ppv);
    return CLASS_E_CLASSNOTAVAILABLE;
}
__declspec(dllexport) HRESULT WINAPI DllCanUnloadNow(void) { return obj_count ? S_FALSE : S_OK; }

BOOL WINAPI DllMain(HINSTANCE inst, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(inst);
    return TRUE;
}
