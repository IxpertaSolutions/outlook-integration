#pragma once

#include <string>
#include <windows.h>
#include <winnt.h>

typedef int (*startConversation_t)(const wchar_t *wstr);
typedef void (*log_t)(
	const wchar_t *file,
	const int line,
	const wchar_t *function_name,
	const wchar_t *msg);
