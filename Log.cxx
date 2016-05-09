/*
 * Outlook Integration library.
 *
 * Copyright (c) 2016, Ixperta Solutions s.r.o.
 *
 * This work is based on
 * Jitsi, the OpenSource Java VoIP and Instant Messaging client.
 *
 * Copyright @ 2015 Atlassian Pty Ltd
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
#include "Log.h"

#include <stdarg.h>
#include <sstream>
#include "StringUtils.h"
#include "OutlookIntegrationInterface.h"

#define MAX_MSG_LENGTH 500

void Log::d(const wchar_t *file,
	const int line,
	const char *function_name,
	LPCTSTR format,
	...)
{
	std::wostringstream wstream;
	wchar_t msg[MAX_MSG_LENGTH];
	std::string func_str(function_name);
	std::wstring func_wstr(func_str.begin(), func_str.end());

	va_list args;
	va_start(args, format);
	vswprintf(msg, MAX_MSG_LENGTH, format, args);
	va_end(args);

	OutlookIntegrationInterface::getInstance()->log(file, line , func_wstr.c_str(), msg);
}
