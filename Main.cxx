#include <windows.h>

#include "OutOfProcessServer.h"
#include "OutlookIntegrationInterface.h"

int WINAPI WinMain(HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPSTR lpCmdLine,
	int nCmdShow)
{
	OutlookIntegrationInterface::getInstance();
	while (1) { Sleep(1000); }
	return 0;
}