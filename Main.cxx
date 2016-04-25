#include <windows.h>

#include "OutOfProcessServer.h"

int WINAPI WinMain(HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPSTR lpCmdLine,
	int nCmdShow)
{
	OutOfProcessServer::start();
	while (1) { Sleep(1000); }
	return 0;
}