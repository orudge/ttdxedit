#include <windows.h>
#include <shellapi.h>

int WINAPI WinMain(
	 HINSTANCE hInstance,	// handle to current instance
	 HINSTANCE hPrevInstance,	// handle to previous instance
	 LPSTR lpCmdLine,	// pointer to command line
	 int nCmdShow 	// show state of window
	)
{
	HINSTANCE hInst;

	hInst = ShellExecute(NULL, "install", lpCmdLine, NULL, NULL, SW_SHOWDEFAULT);

	if (hInst > 32)
		WaitForSingleObject(hInst, INFINITE);

	return(0);
}