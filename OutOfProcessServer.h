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
#ifndef _JMSOFFICECOMM_OUTOFPROCESSSERVER_H_
#define _JMSOFFICECOMM_OUTOFPROCESSSERVER_H_

#include <objbase.h>
#include <windows.h>

class ClassFactory;

/**
 * Implements the jmsoutlookcomm application which is an out-of-process local
 * COM server.
 *
 * @author Lyubomir Marinov
 */
class OutOfProcessServer
{
public:
    static ULONG addRef() { /* TODO Auto-generated method stub */ return 0; }
    static void enterCriticalSection()
        { ::EnterCriticalSection(&_criticalSection); }
    static LPSTR getClassName(LPCSTR className);
    static HANDLE getThreadHandle() { return _threadHandle; }
    static DWORD getThreadId() { return _threadId; }
    static HRESULT getTypeInfo(UINT index, ITypeInfo **ppTInfo)
        { return _iTypeLib->GetTypeInfo(index, ppTInfo); }
    static UINT getTypeInfoCount() { return _iTypeLib->GetTypeInfoCount(); }
    static HRESULT getTypeInfoOfGuid(REFGUID guid, ITypeInfo **ppTInfo)
        { return _iTypeLib->GetTypeInfoOfGuid(guid, ppTInfo); }
    static void leaveCriticalSection()
        { ::LeaveCriticalSection(&_criticalSection); }
    static ULONG release() { /* TODO Auto-generated method stub */ return 0; }
    static HRESULT start();
    static HRESULT stop();

private:
    static BOOL isMicrosoftOfficeOutlookCallIntegrationIMApplication();
    static HRESULT loadRegTypeLib();
    static DWORD regCreateKeyAndSetValue(LPCTSTR key, LPCTSTR valueName, DWORD data);
    static HRESULT registerClassObjects();
    static ULONG releaseTypeLib();
    static HRESULT revokeClassObjects();
    static unsigned __stdcall run(void *);
    static DWORD setIMProvidersCommunicatorUpAndRunning(DWORD dw);

    static CRITICAL_SECTION _criticalSection;
    static LPTYPELIB        _iTypeLib;
    static ClassFactory *   _messengerClassFactory;
    static HANDLE           _threadHandle;
    static DWORD            _threadId;
};

#endif /* #ifndef _JMSOFFICECOMM_OUTOFPROCESSSERVER_H_ */
