#include <windows.h>
#include <time.h>

BOOL WINAPI DllMain(
  HINSTANCE hinstDLL,
  DWORD fdwReason,
  LPVOID lpvReserved
)
{
	return(TRUE);
}

int WINAPI __export GetUNIXTime()
{
	return(time(NULL));
}

typedef BOOL (WINAPI *def_AllocateAndInitializeSid)(
  __in   PSID_IDENTIFIER_AUTHORITY pIdentifierAuthority,
  __in   BYTE nSubAuthorityCount,
  __in   DWORD dwSubAuthority0,
  __in   DWORD dwSubAuthority1,
  __in   DWORD dwSubAuthority2,
  __in   DWORD dwSubAuthority3,
  __in   DWORD dwSubAuthority4,
  __in   DWORD dwSubAuthority5,
  __in   DWORD dwSubAuthority6,
  __in   DWORD dwSubAuthority7,
  __out  PSID *pSid
);

typedef BOOL (WINAPI *def_CheckTokenMembership)(
  __in_opt  HANDLE TokenHandle,
  __in      PSID SidToCheck,
  __out     PBOOL IsMember
);

typedef PVOID (WINAPI *def_FreeSid)(
  __in  PSID pSid
);

int WINAPI __export IsUserAnAdmin()
{
	SID_IDENTIFIER_AUTHORITY NTAuth = SECURITY_NT_AUTHORITY;
	HMODULE hAdvApi32;
	int RetCode = 0;
	PSID AdminSID;

	def_AllocateAndInitializeSid p_AllocateAndInitializeSid;
	def_CheckTokenMembership p_CheckTokenMembership;
	def_FreeSid p_FreeSid;

	hAdvApi32 = LoadLibrary("advapi32.dll");

	if (!hAdvApi32)
		return 1;

	p_AllocateAndInitializeSid = GetProcAddress(hAdvApi32, "AllocateAndInitializeSid");
	p_CheckTokenMembership = GetProcAddress(hAdvApi32, "CheckTokenMembership");
	p_FreeSid = (def_FreeSid) GetProcAddress(hAdvApi32, "FreeSid");

	/* Check if we're running on NT */
	if ((!p_AllocateAndInitializeSid) || (!p_CheckTokenMembership) || (!p_FreeSid))
	{
		FreeLibrary(hAdvApi32);
		return 1;
	}
	
	if (p_AllocateAndInitializeSid(&NTAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &AdminSID))
	{
		if (!p_CheckTokenMembership(NULL, AdminSID, &RetCode))
			RetCode = 0;

		p_FreeSid(AdminSID);
	}

	FreeLibrary(hAdvApi32);
	return RetCode;
}

