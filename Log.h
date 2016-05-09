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
#ifndef _JMSOFFICECOMM_LOG_H_
#define _JMSOFFICECOMM_LOG_H_

#include <stdio.h>
#include <tchar.h>
#include <windows.h>

#define DEBUG(format, ...) Log::d(WIDEN(__FILE__), __LINE__, __func__, format, __VA_ARGS__)
#define WIDEN2(x) L ## x
#define WIDEN(x) WIDEN2(x)

class Log
{
public:

    static void d(const wchar_t *file,
		const int line,
		const char *function_name,
		LPCTSTR format,
		...);
};

#endif /* #ifndef _JMSOFFICECOMM_LOG_H_ */
