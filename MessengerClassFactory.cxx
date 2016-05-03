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
#include "MessengerClassFactory.h"

#include "Messenger.h"

 // {F03BE5F9-058B-4C6A-9819-0AF3849B36A9}
EXTERN_C const GUID DECLSPEC_SELECTANY CLSID_Messenger
    = { 0xf03be5f9, 0x58b, 0x4c6a,{ 0x98, 0x19, 0xa, 0xf3, 0x84, 0x9b, 0x36, 0xa9 } };


STDMETHODIMP
MessengerClassFactory::CreateInstance(LPUNKNOWN outer, REFIID iid, PVOID *obj)
{
    HRESULT hr;

    if (outer)
    {
        *obj = NULL;
        hr = CLASS_E_NOAGGREGATION;
    }
    else
    {
        IMessenger *messenger;

        if (_messenger)
        {
            hr = _messenger->Resolve(IID_IMessenger, (PVOID *) &messenger);
            if (FAILED(hr) && (E_NOINTERFACE != hr))
            {
                _messenger->Release();
                _messenger = NULL;
            }
        }
        else
            messenger = NULL;
        if (!messenger)
        {
            messenger = new Messenger();

            IWeakReferenceSource *weakReferenceSource;

            hr
                = messenger->QueryInterface(
                        IID_IWeakReferenceSource,
                        (PVOID *) &weakReferenceSource);
            if (SUCCEEDED(hr))
            {
                IWeakReference *weakReference;

                hr = weakReferenceSource->GetWeakReference(&weakReference);
                if (SUCCEEDED(hr))
                {
                    if (_messenger)
                        _messenger->Release();
                    _messenger = weakReference;
                }
            }
        }
        hr = messenger->QueryInterface(iid, obj);
        messenger->Release();
    }
    return hr;
}
