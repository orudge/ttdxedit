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

int WINAPI GetUNIXTime()
{
	return(time(NULL));
}
