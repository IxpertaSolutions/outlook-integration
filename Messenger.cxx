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
#include "Messenger.h"

#include "ConnectionPoint.h"
#include "MessengerContact.h"
#include "MessengerContacts.h"
#include "MessengerServices.h"
#include "OutlookIntegrationInterface.h"
#include <olectl.h>

EXTERN_C const GUID DECLSPEC_SELECTANY DIID_DMessengerEvents
    = { 0xC9A6A6B6, 0x9BC1, 0x43a5, { 0xB0, 0x6B, 0xE5, 0x88, 0x74, 0xEE, 0xBC, 0x96 } };

EXTERN_C const GUID DECLSPEC_SELECTANY IID_ICallFactory
    = { 0x1c733a30, 0x2a1c, 0x11ce, { 0xad, 0xe5, 0x00, 0xaa, 0x00, 0x44, 0x77, 0x3d } };

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IMessenger
    = { 0xD50C3186, 0x0F89, 0x48f8, { 0xB2, 0x04, 0x36, 0x04, 0x62, 0x9D, 0xEE, 0x10 } };

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IMessenger2
    = { 0xD50C3286, 0x0F89, 0x48f8, { 0xB2, 0x04, 0x36, 0x04, 0x62, 0x9D, 0xEE, 0x10 } };

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IMessenger3
    = { 0xD50C3386, 0x0F89, 0x48f8, { 0xB2, 0x04, 0x36, 0x04, 0x62, 0x9D, 0xEE, 0x10 } };

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IMessengerAdvanced
    = { 0xDA0635E8, 0x09AF, 0x480c, { 0x88, 0xB2, 0xAA, 0x9F, 0xA1, 0xD9, 0xDB, 0x27 } };

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IMessengerContactResolution
    = { 0x53A5023D, 0x6872, 0x454a, { 0x9A, 0x4F, 0x82, 0x7F, 0x18, 0xCF, 0xBE, 0x02 } };

class OnContactStatusChangeEvent
{
public:
    OnContactStatusChangeEvent(BSTR signinName, MISTATUS status)
        : _signinName(signinName), _status(status) {}
    ~OnContactStatusChangeEvent()
        { ::SysFreeString(_signinName); }

    BSTR _signinName;
    MISTATUS _status;
};

Messenger *Messenger::_singleton = NULL;

Messenger::Messenger()
    : _dMessengerEventsConnectionPoint(NULL),
      _messengerContactCount(0),
      _messengerContacts(NULL),
      _myContacts(NULL),
      _services(NULL)
{
    if (SUCCEEDED(::StringFromCLSID(CLSID_Messenger, &_myServiceId)))
        {}
    else
        _myServiceId = NULL;

    _singleton = this;
}

Messenger::~Messenger()
{
    if (_singleton == this)
        _singleton = NULL;

    if (_dMessengerEventsConnectionPoint)
        delete _dMessengerEventsConnectionPoint;

    // _messengerContacts
    if (_messengerContactCount)
    {
        size_t i = 0;
        IWeakReference **messengerContactIt = _messengerContacts;

        for (; i < _messengerContactCount; i++, messengerContactIt++)
            (*messengerContactIt)->Release();
        ::free(_messengerContacts);
    }

    if (_myContacts)
        _myContacts->Release();
    if (_myServiceId)
        ::CoTaskMemFree(_myServiceId);
    if (_services)
        _services->Release();
}

STDMETHODIMP Messenger::AddContact(long hwndParent, BSTR bstrEMail)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::AutoSignin()
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::CreateGroup(BSTR bstrName, VARIANT vService, IDispatch **ppGroup)
    STDMETHODIMP_E_NOTIMPL_STUB

HRESULT
Messenger::createMessengerContact(BSTR signinName, REFIID iid, PVOID *obj)
{
    MessengerContact *messengerContact = new MessengerContact(this, signinName);
    HRESULT hr;

    if (messengerContact)
    {
        hr = messengerContact->QueryInterface(iid, obj);

        /*
         * We've created a new instance and we've asked it about the requested
         * interface. What follows is keeping track of the instance for the
         * specified signin name in order to try to avoid having multiple
         * instances for one and the same signin name at one and the same time
         * (because MSDN mentions it) and it is not vital.
         */
        if (SUCCEEDED(hr))
        {
            IWeakReferenceSource *weakReferenceSource;

            if (SUCCEEDED(
                    messengerContact->QueryInterface(
                            IID_IWeakReferenceSource,
                            (PVOID *) &weakReferenceSource)))
            {
                IWeakReference *weakReference;

                if (SUCCEEDED(
                        weakReferenceSource->GetWeakReference(&weakReference)))
                {
                    size_t newMessengerContactCount = _messengerContactCount + 1;
                    IWeakReference **newMessengerContacts
                        = (IWeakReference **)
                            ::realloc(
                                    _messengerContacts,
                                    newMessengerContactCount
                                        * sizeof(IWeakReference *));

                    if (newMessengerContacts)
                    {
                        newMessengerContacts[newMessengerContactCount - 1]
                            = weakReference;
                        _messengerContactCount = newMessengerContactCount;
                        _messengerContacts = newMessengerContacts;
                    }
                    else
                        weakReference->Release();
                }
                weakReferenceSource->Release();
            }
        }

        messengerContact->Release();
    }
    else
    {
        *obj = NULL;
        hr = E_OUTOFMEMORY;
    }
    return hr;
}

STDMETHODIMP Messenger::EnumConnectionPoints(IEnumConnectionPoints **ppEnum)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP)
{
    HRESULT hr;

    if (ppCP)
    {
        if (DIID_DMessengerEvents == riid)
        {
            if (!_dMessengerEventsConnectionPoint)
            {
                _dMessengerEventsConnectionPoint
                    = new DMessengerEventsConnectionPoint(this);
            }
            _dMessengerEventsConnectionPoint->AddRef();
            *ppCP = _dMessengerEventsConnectionPoint;
            hr = S_OK;
        }
        else
        {
            *ppCP = NULL;
            hr = CONNECT_E_NOCONNECTION;
        }
    }
    else
        hr = E_POINTER;
    return hr;
}

STDMETHODIMP Messenger::FindContact(long hwndParent, BSTR bstrFirstName, BSTR bstrLastName, VARIANT vbstrCity, VARIANT vbstrState, VARIANT vbstrCountry)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_ContactsSortOrder(MUASORT *pSort)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_MyContacts(IDispatch **ppMContacts)
    STDMETHODIMP_RESOLVE_WEAKREFERENCE_OR_NEW(ppMContacts,_myContacts,MessengerContacts,this)

STDMETHODIMP Messenger::get_MyFriendlyName(BSTR *pbstrName)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_MyGroups(IDispatch **ppMGroups)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_MyPhoneNumber(MPHONE_TYPE PhoneType, BSTR *pbstrNumber)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_MyProperty(MCONTACTPROPERTY ePropType, VARIANT *pvPropVal)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_MyServiceId(BSTR *pbstrServiceId)
{
    HRESULT hr;

    if (pbstrServiceId)
    {
        if (_myServiceId)
        {
            hr
                = ((*pbstrServiceId = ::SysAllocString(_myServiceId)))
                    ? S_OK
                    : E_OUTOFMEMORY;
        }
        else
        {
            *pbstrServiceId = NULL;
            hr = E_FAIL;
        }
    }
    else
        hr = E_INVALIDARG;
    return hr;
}

STDMETHODIMP Messenger::get_MyServiceName(BSTR *pbstrServiceName)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_MySigninName(BSTR *pbstrName)
{
    return get_MyServiceId(pbstrName);
}

STDMETHODIMP Messenger::get_MyStatus(MISTATUS *pmStatus)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_Property(MMESSENGERPROPERTY ePropType, VARIANT *pvPropVal)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_ReceiveFileDirectory(BSTR *bstrPath)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_Services(IDispatch **ppdispServices)
    STDMETHODIMP_RESOLVE_WEAKREFERENCE_OR_NEW(ppdispServices,_services,MessengerServices,this)

STDMETHODIMP Messenger::get_UnreadEmailCount(MUAFOLDER mFolder, LONG *plCount)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::get_Window(IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::GetAuthenticationInfo(BSTR *pbstrAuthInfo)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP
Messenger::GetContact(BSTR bstrSigninName, BSTR bstrServiceId, IDispatch **ppMContact)
{
    HRESULT hr;

    if (ppMContact)
    {
        if (bstrSigninName)
        {
            /*
             * Try to find an existing MessengerContact instance which has not
             * been released to deletion yet and which has the specified signin
             * name.
             */
            hr
                = getMessengerContact(
                        bstrSigninName,
                        IID_IDispatch,
                        (PVOID *) ppMContact);
            if (FAILED(hr))
            {
                /*
                 * Try to find a contact with the specified signin name in the
                 * MyContacts collection of this Messenger.
                 */
                LPDISPATCH iDispatch;

                hr = get_MyContacts(&iDispatch);
                if (SUCCEEDED(hr))
                {
                    IMessengerContacts *myContacts;

                    hr
                        = iDispatch->QueryInterface(
                                IID_IMessengerContacts,
                                (PVOID *) &myContacts);
                    iDispatch->Release();
                    if (SUCCEEDED(hr))
                    {
                        LONG myContactCount;

                        hr = myContacts->get_Count(&myContactCount);
                        if (SUCCEEDED(hr))
                        {
                            for (LONG i = 0; i < myContactCount; i++)
                            {
                                hr = myContacts->Item(i, &iDispatch);
                                if (SUCCEEDED(hr))
                                {
                                    if (MessengerContact::signinNameEquals(
                                            iDispatch,
                                            bstrSigninName))
                                    {
                                        *ppMContact = iDispatch;
                                        break;
                                    }
                                    else
                                        iDispatch->Release();
                                }
                                else
                                    break;
                            }
                        }
                        myContacts->Release();
                    }
                }

                if (FAILED(hr) || !(*ppMContact))
                {
                    hr
                        = createMessengerContact(
                                bstrSigninName,
                                IID_IDispatch,
                                (PVOID *) ppMContact);
                }
            }
        }
        else
        {
            *ppMContact = NULL;
            hr = E_FAIL;
        }
    }
    else
        hr = RPC_X_NULL_REF_POINTER;
    return hr;
}

HRESULT Messenger::getMessengerContact(BSTR signinName, REFIID iid, PVOID *obj)
{
    *obj = NULL;

    size_t i = 0;
    IWeakReference **messengerContactIt = _messengerContacts;
    HRESULT hr = E_FAIL;

    while (i < _messengerContactCount)
    {
        IWeakReference *weakReference = *messengerContactIt;
        LPDISPATCH iDispatch;

        hr = weakReference->Resolve(IID_IDispatch, (PVOID *) &iDispatch);
        if (SUCCEEDED(hr))
        {
            if (MessengerContact::signinNameEquals(iDispatch, signinName))
            {
                if (IID_IDispatch == iid)
                    *obj = iDispatch;
                else
                {
                    hr = weakReference->Resolve(iid, obj);
                    iDispatch->Release();
                }
                break;
            }
            else
                iDispatch->Release();

            i++;
            messengerContactIt++;
        }
        else if (E_NOINTERFACE != hr)
        {
            /*
             * The weakReference appears to have been invalidated. Release the
             *
			 s associated with it.
             */
            *messengerContactIt = NULL;
            weakReference->Release();

            _messengerContactCount--;

            /*
             * Move the emptied slot of the _messengerContacts storage at the
             * end where it is not accessible given the value of
             * _messengerContactCount.
             */
            size_t j = i;
            IWeakReference **it = messengerContactIt;

            for (; j < _messengerContactCount; j++)
            {
                IWeakReference **nextIt = it + 1;

                *it = *nextIt;
                it = nextIt;
            }
        }
    }

    if (SUCCEEDED(hr) && !(*obj))
        hr = E_FAIL;
    return hr;
}

STDMETHODIMP Messenger::InstantMessage(VARIANT vContact, IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::InviteApp(VARIANT vContact, BSTR bstrAppID, IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::MediaWizard(long hwndParent)
    STDMETHODIMP_E_NOTIMPL_STUB

void CALLBACK Messenger::onContactStatusChange(ULONG_PTR dwParam)
{
    BOOL run;

    OutOfProcessServer::enterCriticalSection();

    HANDLE threadHandle = OutOfProcessServer::getThreadHandle();

    if (threadHandle)
    {
        run = (::GetCurrentThreadId() == OutOfProcessServer::getThreadId());
        if (!run
                && !::QueueUserAPC(
                        onContactStatusChange,
                        threadHandle,
                        dwParam))
        {
            delete (OnContactStatusChangeEvent *) dwParam;
        }
    }
    else
    {
        run = FALSE;
        delete (OnContactStatusChangeEvent *) dwParam;
    }

    OutOfProcessServer::leaveCriticalSection();

    if (run)
    {
        Messenger *thiz = _singleton;
        OnContactStatusChangeEvent *event
            = (OnContactStatusChangeEvent *) dwParam;

        if (thiz && thiz->_dMessengerEventsConnectionPoint)
        {
            HRESULT hr;
            LPDISPATCH contact;

            hr = thiz->GetContact(event->_signinName, NULL, &contact);
            if (SUCCEEDED(hr))
            {
                hr
                    = thiz->_dMessengerEventsConnectionPoint
                            ->OnContactStatusChange(contact, event->_status);
                contact->Release();
            }
        }

        delete event;
    }
}

STDMETHODIMP Messenger::OpenInbox()
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::OptionsPages(long hwndParent, MOPTIONPAGE mOptionPage)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::Page(VARIANT vContact, IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::Phone(VARIANT vContact, MPHONE_TYPE ePhoneNumber, BSTR bstrNumber, IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::put_ContactsSortOrder(MUASORT Sort)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::put_MyProperty(MCONTACTPROPERTY ePropType, VARIANT vPropVal)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::put_MyStatus(MISTATUS mStatus)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::put_Property(MMESSENGERPROPERTY ePropType, VARIANT vPropVal)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::QueryInterface(REFIID iid, PVOID *obj)
{
    HRESULT hr;

    if (!obj)
        hr = E_POINTER;
    else if (IID_IMessenger == iid)
    {
        AddRef();
        *obj = static_cast<IMessenger *>(this);
        hr = S_OK;
    }
    else if (IID_IMessenger2 == iid)
    {
        AddRef();
        *obj = static_cast<IMessenger2 *>(this);
        hr = S_OK;
    }
    else if (IID_IMessenger3 == iid)
    {
        AddRef();
        *obj = static_cast<IMessenger3 *>(this);
        hr = S_OK;
    }
    else if (IID_IMessengerAdvanced == iid)
    {
        AddRef();
        *obj = static_cast<IMessengerAdvanced *>(this);
        hr = S_OK;
    }
    else if (IID_IConnectionPointContainer == iid)
    {
        AddRef();
        *obj = static_cast<IConnectionPointContainer *>(this);
        hr = S_OK;
    }
    else if (IID_IMessengerContactResolution == iid)
    {
        AddRef();
        *obj = static_cast<IMessengerContactResolution *>(this);
        hr = S_OK;
    }
    else
        hr = DispatchImpl::QueryInterface(iid, obj);
    return hr;
}

STDMETHODIMP
Messenger::ResolveContact(ADDRESS_TYPE AddressType, CONTACT_RESOLUTION_TYPE ResolutionType, BSTR bstrAddress, BSTR *pbstrIMAddress)
{
    HRESULT hr;

    if (pbstrIMAddress)
    {
        if (bstrAddress)
        {
            if (ADDRESS_TYPE_SMTP == AddressType)
            {
                hr
                    = ((*pbstrIMAddress = ::SysAllocString(bstrAddress)))
                        ? S_OK
                        : E_OUTOFMEMORY;
            }
            else
            {
                *pbstrIMAddress = NULL;
                hr = E_NOTIMPL;
            }
        }
        else
        {
            *pbstrIMAddress = NULL;
            hr = E_INVALIDARG;
        }
    }
    else
        hr = RPC_X_NULL_REF_POINTER;
    return hr;
}

STDMETHODIMP Messenger::SendFile(VARIANT vContact, BSTR bstrFileName, IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::SendMail(VARIANT vContact)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::Signin(long hwndParent, BSTR bstrSigninName, BSTR bstrPassword)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::Signout()
    STDMETHODIMP_E_NOTIMPL_STUB

HRESULT Messenger::start()
{
    return S_OK;
}

HRESULT Messenger::toString(CComVariant &v, std::wstring *wstr)
{
	BSTR bstr = NULL;
	HRESULT hr;

	if (VT_BSTR == v.vt)
	{
		bstr = v.bstrVal;
		hr = S_OK;
	}
	else if ((VT_BSTR | VT_BYREF) == v.vt)
	{
		bstr = *(v.pbstrVal);
		hr = S_OK;
	}
	else
		hr = E_INVALIDARG;

	if (SUCCEEDED(hr))
	{
		if (bstr)
		{
			*wstr = std::wstring(bstr);
		}
		else
		{
			wstr = NULL;
			hr = E_INVALIDARG;
		}
	}

	return hr;
}

HRESULT
Messenger::toStringArray(VARIANT &v, std::list<std::wstring> *wstringArray)
{
	HRESULT hr;

	if (VT_ARRAY == (v.vt & VT_ARRAY))
	{
		SAFEARRAY *sa = v.parray;

		if (sa
			&& (1 == sa->cDims)
			&& (FADF_VARIANT == (sa->fFeatures & FADF_VARIANT)))
		{
			CComSafeArray<VARIANT> superSa;
			superSa.Attach(sa);
			for (ULONG i = 0; i < superSa.GetCount(); i++)
			{
				std::wstring _wstring;

				VARIANT obj = superSa.GetAt(i);
				hr = toString(CComVariant(obj), &_wstring);
				if (SUCCEEDED(hr))
				{
					wstringArray->push_back(_wstring);
				}
			}
			superSa.Detach();
		}
		else
			hr = E_INVALIDARG;
	}
	else
		hr = E_INVALIDARG;
	return hr;
}

STDMETHODIMP
Messenger::StartConversation
	(CONVERSATION_TYPE conversationType,
	VARIANT vParticipants,
	VARIANT vContextualData,
	VARIANT vSubject,
	VARIANT vConversationIndex,
	VARIANT vConversationData,
	VARIANT *pvWndHnd)
{
	Log::d(_T("%s\n"), FUNC_NAME_MACRO);
	Log::d(_T("%d\n"), conversationType);
	if (VT_ARRAY == (vParticipants.vt & VT_ARRAY))
	{
		Log::d(_T("looks like safe_array\n"));
	}

	std::list<std::wstring> wstringList;

	HRESULT hr = toStringArray(vParticipants, &wstringList);
	if (SUCCEEDED(hr))
	{
		Log::d(_T("toStringArray ok\n"));
	}
	else
	{
		Log::d(_T("toStringArray error: %d\n"), hr);
	}

	for (std::wstring & wstr : wstringList)
	{
		std::string str(wstr.begin(), wstr.end());
		Log::d(_T("%s"), str.c_str());
	}

	std::wstring wstr;
    hr = toString(CComVariant(vConversationData), &wstr);
	if (SUCCEEDED(hr))
	{
		Log::d(_T("toString ok\n"));
	}
	else
	{
		Log::d(_T("toString error: %d\n"), hr);
		return E_INVALIDARG;
	}
	std::string str(wstr.begin(), wstr.end());
	Log::d(_T("%s"), wstr.c_str());

	return OutlookIntegrationInterface::getInstance()->callStartConversation(wstr.c_str());
}

STDMETHODIMP Messenger::StartVideo(VARIANT vContact, IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP Messenger::StartVoice(VARIANT vContact, IDispatch **ppMWindow)
    STDMETHODIMP_E_NOTIMPL_STUB

HRESULT Messenger::stop()
{
    return S_OK;
}

STDMETHODIMP Messenger::ViewProfile(VARIANT vContact)
    STDMETHODIMP_E_NOTIMPL_STUB
