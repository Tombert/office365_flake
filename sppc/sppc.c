/*
 * sppc.dll shim for running Microsoft Office click-to-run under Wine/Proton.
 *
 * Wine's own sppc.dll (Software Protection Platform client) is almost entirely
 * stubs, and calling a stub aborts the process. The Office integrator calls
 * SLInstallLicense while installing licences and dies with 0x80000100, which
 * the click-to-run bootstrapper reports as error 0-2031 (17002).
 *
 * This shim exports every entry point of the real sppc.dll and returns benign
 * values: install/uninstall style calls succeed, queries report "not found" so
 * callers fall through to their non-SPP paths (Office 365 uses vNext licensing
 * once signed in, which lives in the user profile, not SPP).
 *
 * x86_64 only: the Win64 ABI is caller-clean, so a wrong parameter count is
 * harmless. Do not build this for i386 without giving every function its exact
 * stdcall signature.
 */
#include <windows.h>
#include <string.h>

typedef HANDLE HSLC;
typedef HANDLE HSLP;
typedef GUID   SLID;

#ifndef SL_E_VALUE_NOT_FOUND
#define SL_E_VALUE_NOT_FOUND       ((HRESULT)0xC004F012)
#define SL_E_RIGHT_NOT_CONSUMED    ((HRESULT)0xC004F013)
#define SL_E_RIGHT_NOT_GRANTED     ((HRESULT)0xC004F00C)
#endif

#define API __declspec(dllexport) HRESULT WINAPI

static void fake_guid(SLID *id)
{
    static LONG counter;
    if (!id) return;
    memset(id, 0, sizeof(*id));
    id->Data1 = 0x5da2c9b4;
    id->Data2 = 0x1c2f;
    id->Data3 = 0x4c1e;
    id->Data4[0] = 0x9a; id->Data4[1] = 0x7e;
    *(LONG *)&id->Data4[4] = InterlockedIncrement(&counter);
}

/* ---- session ---------------------------------------------------------- */
API SLOpen(HSLC *handle)
{
    if (!handle) return E_INVALIDARG;
    *handle = (HSLC)(ULONG_PTR)0x534c4f50; /* 'SLOP' */
    return S_OK;
}
API SLClose(HSLC handle) { (void)handle; return S_OK; }

/* ---- licence / proof-of-purchase installation: pretend it worked ------- */
API SLInstallLicense(HSLC h, UINT cb, const BYTE *blob, SLID *file_id)
{
    (void)h; (void)cb; (void)blob;
    fake_guid(file_id);
    return S_OK;
}
API SLUninstallLicense(HSLC h, const SLID *id) { (void)h; (void)id; return S_OK; }

API SLInstallProofOfPurchase(HSLC h, PCWSTR alg, PCWSTR key, UINT cb, PBYTE data, SLID *pkey_id)
{
    (void)h; (void)alg; (void)key; (void)cb; (void)data;
    fake_guid(pkey_id);
    return S_OK;
}
API SLInstallProofOfPurchaseEx(HSLC h, const SLID *app, PCWSTR alg, PCWSTR key, UINT cb, PBYTE data, SLID *pkey_id)
{
    (void)h; (void)app; (void)alg; (void)key; (void)cb; (void)data;
    fake_guid(pkey_id);
    return S_OK;
}
API SLUninstallProofOfPurchase(HSLC h, const SLID *id) { (void)h; (void)id; return S_OK; }
API SLSetCurrentProductKey(HSLC h, const SLID *product, const SLID *pkey) { (void)h; (void)product; (void)pkey; return S_OK; }

/* ---- queries: nothing is ever found ---------------------------------- */
API SLGetLicensingStatusInformation(HSLC h, const SLID *app, const SLID *product, PCWSTR right, UINT *count, void **status)
{
    (void)h; (void)app; (void)product; (void)right;
    if (count) *count = 0;
    if (status) *status = NULL;
    return SL_E_RIGHT_NOT_CONSUMED; /* same as Wine's builtin */
}
API SLGetSLIDList(HSLC h, int qtype, const SLID *qid, int rtype, UINT *count, SLID **ids)
{
    (void)h; (void)qtype; (void)qid; (void)rtype;
    if (count) *count = 0;
    if (ids) *ids = NULL;
    return S_OK;
}
API SLGetInstalledProductKeyIds(HSLC h, const SLID *product, UINT *count, SLID **ids)
{
    (void)h; (void)product;
    if (count) *count = 0;
    if (ids) *ids = NULL;
    return S_OK;
}

/* generic "name -> typed blob" getters: (h, id, name, type*, size*, data**) */
#define NOT_FOUND_GETTER(name) \
    API name(HSLC h, const SLID *id, PCWSTR value, int *type, UINT *size, PBYTE *data) \
    { (void)h; (void)id; (void)value; if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }
NOT_FOUND_GETTER(SLGetPKeyInformation)
NOT_FOUND_GETTER(SLGetProductSkuInformation)
NOT_FOUND_GETTER(SLGetLicenseInformation)
NOT_FOUND_GETTER(SLGetApplicationInformation)

/* policy getters take a policy handle and no SLID: (h, name, type*, size*, data**) */
API SLGetPolicyInformation(HSLP h, PCWSTR value, int *type, UINT *size, PBYTE *data)
{ (void)h; (void)value; if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGetApplicationPolicy(HSLP h, PCWSTR value, int *type, UINT *size, PBYTE *data)
{ (void)h; (void)value; if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }

API SLGetServiceInformation(HSLC h, PCWSTR value, int *type, UINT *size, PBYTE *data)
{ (void)h; (void)value; if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGetPolicyInformationDWORD(HSLP h, PCWSTR value, DWORD *out)
{ (void)h; (void)value; if (out) *out = 0; return SL_E_VALUE_NOT_FOUND; }
API SLGetGenuineInformation(const SLID *id, PCWSTR value, int *type, UINT *size, PBYTE *data)
{ (void)id; (void)value; if (type) *type = 0; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGetLicense(HSLC h, const SLID *file_id, UINT *size, PBYTE *data)
{ (void)h; (void)file_id; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGetLicenseFileId(HSLC h, UINT cb, const BYTE *blob, SLID *file_id)
{ (void)h; (void)cb; (void)blob; if (file_id) memset(file_id, 0, sizeof(*file_id)); return SL_E_VALUE_NOT_FOUND; }
API SLGetPKeyId(HSLC h, PCWSTR alg, PCWSTR key, UINT cb, const BYTE *data, SLID *pkey_id)
{ (void)h; (void)alg; (void)key; (void)cb; (void)data; if (pkey_id) memset(pkey_id, 0, sizeof(*pkey_id)); return SL_E_VALUE_NOT_FOUND; }
API SLGetActiveLicenseInfo(HSLC h, const SLID *app, const SLID *product, UINT *count, void **info)
{ (void)h; (void)app; (void)product; if (count) *count = 0; if (info) *info = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGetEncryptedPIDEx(HSLC h, const SLID *product, UINT *size, PBYTE *data)
{ (void)h; (void)product; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGetAuthenticationResult(HSLC h, UINT *size, PBYTE *data)
{ (void)h; if (size) *size = 0; if (data) *data = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGenerateOfflineInstallationId(HSLC h, const SLID *product, PWSTR *out)
{ (void)h; (void)product; if (out) *out = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGenerateOfflineInstallationIdEx(HSLC h, const SLID *product, const void *info, PWSTR *out)
{ (void)h; (void)product; (void)info; if (out) *out = NULL; return SL_E_VALUE_NOT_FOUND; }
API SLGatherMigrationBlob(BOOL migratable, BOOL user, UINT *size, PBYTE data)
{ (void)migratable; (void)user; (void)data; if (size) *size = 0; return S_OK; }
API SLGatherMigrationBlobEx(BOOL migratable, BOOL user, UINT *size, PBYTE data)
{ (void)migratable; (void)user; (void)data; if (size) *size = 0; return S_OK; }

/* rights: never granted through SPP, so Office asks its own (vNext) licensing */
API SLConsumeRight(HSLC h, const SLID *app, const SLID *product, PCWSTR right, void *reserved)
{ (void)h; (void)app; (void)product; (void)right; (void)reserved; return SL_E_RIGHT_NOT_GRANTED; }

API SLIsGenuineLocalEx(const SLID *app, const SLID *alt, int *state)
{ (void)app; (void)alt; if (state) *state = 0 /* SL_GEN_STATE_IS_GENUINE */; return S_OK; }

API SLLoadApplicationPolicies(const SLID *app, const SLID *product, DWORD flags, HSLP *handle)
{ (void)app; (void)product; (void)flags; if (handle) *handle = (HSLP)(ULONG_PTR)0x534c504f; return S_OK; }
API SLUnloadApplicationPolicies(HSLP handle, DWORD flags) { (void)handle; (void)flags; return S_OK; }
API SLPersistApplicationPolicies(const SLID *app, const SLID *product, DWORD flags)
{ (void)app; (void)product; (void)flags; return S_OK; }
API SLRegisterEvent(HSLC h, PCWSTR name, const SLID *app, HANDLE event) { (void)h; (void)name; (void)app; (void)event; return S_OK; }
API SLUnregisterEvent(HSLC h, PCWSTR name, const SLID *app, HANDLE event) { (void)h; (void)name; (void)app; (void)event; return S_OK; }

/* everything else: accept and do nothing */
#define OK_STUB(name) API name(void) { return S_OK; }
OK_STUB(SLCallServer)
OK_STUB(SLDepositMigrationBlob)
OK_STUB(SLDepositOfflineConfirmationId)
OK_STUB(SLDepositOfflineConfirmationIdEx)
OK_STUB(SLDepositStoreToken)
OK_STUB(SLFireEvent)
OK_STUB(SLPersistRTSPayloadOverride)
OK_STUB(SLReArm)
OK_STUB(SLRegisterPlugin)
OK_STUB(SLSetAuthenticationData)
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
    (void)inst; (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(inst);
    return TRUE;
}
