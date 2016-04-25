/*
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
#include "MessengerContact.h"

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IMessengerContact
    = { 0xE7479A0F, 0xBB19, 0x44a5, { 0x96, 0x8F, 0x6F, 0x41, 0xD9, 0x3E, 0xE0, 0xBC } };

EXTERN_C const GUID DECLSPEC_SELECTANY IID_IMessengerContactAdvanced
    = { 0x086F69C0, 0x2FBD, 0x46b3, { 0xBE, 0x50, 0xEC, 0x40, 0x1A, 0xB8, 0x60, 0x99 } };

MessengerContact::MessengerContact(IMessenger *messenger, LPCOLESTR signinName)
    : _messenger(messenger)
{
    _messenger->AddRef();
    if (signinName)
    {
        _signinName = ::_wcsdup(signinName);
    }
    else
        _signinName = NULL;
}

MessengerContact::~MessengerContact()
{
    _messenger->Release();
    if (_signinName)
        ::free(_signinName);
}

STDMETHODIMP MessengerContact::get_Blocked(VARIANT_BOOL *pBoolBlock)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::get_CanPage(VARIANT_BOOL *pBoolPage)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::get_FriendlyName(BSTR *pbstrFriendlyName)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::get_IsSelf(VARIANT_BOOL *pBoolSelf)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::get_IsTagged(VARIANT_BOOL *pBoolIsTagged)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::get_PhoneNumber(MPHONE_TYPE PhoneType, BSTR *bstrNumber)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP
MessengerContact::get_PresenceProperties(VARIANT *pvPresenceProperties)
{
    if (!pvPresenceProperties)
    {
        return E_INVALIDARG;
    }

    MISTATUS status;
    HRESULT hr = get_Status(&status);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = (VT_EMPTY == pvPresenceProperties->vt)
            ? S_OK
            : (::VariantClear(pvPresenceProperties));
    if (FAILED(hr))
    {
        return hr;
    }

    SAFEARRAY *sa = ::SafeArrayCreateVector(VT_VARIANT, 0, PRESENCE_PROP_MAX);
    if (!sa)
    {
        return E_FAIL;
    }

    LONG mstateIndex = PRESENCE_PROP_MSTATE;
    VARIANT vtMState;
    ::VariantInit(&vtMState);
    vtMState.vt = VT_I4;
    vtMState.lVal = status;
    hr = ::SafeArrayPutElement(sa, &mstateIndex, &vtMState);
    if (SUCCEEDED(hr))
    {
        LONG availability;
        switch (status)
        {
            case MISTATUS_AWAY:
            case MISTATUS_OUT_OF_OFFICE:
                availability = 15000;
                break;
            case MISTATUS_BE_RIGHT_BACK:
                availability = 12000;
                break;
            case MISTATUS_BUSY:
            case MISTATUS_IN_A_CONFERENCE:
            case MISTATUS_ON_THE_PHONE:
                availability = 6000;
                break;
            case MISTATUS_DO_NOT_DISTURB:
            case MISTATUS_ALLOW_URGENT_INTERRUPTIONS:
                availability = 9000;
                break;
            case MISTATUS_INVISIBLE:
                availability = 18000;
                break;
            case MISTATUS_ONLINE:
                availability = 3000;
                break;
            default:
                availability = 0;
                break;
        }

        if (availability)
        {
            LONG availIndex = PRESENCE_PROP_AVAILABILITY;
            VARIANT vtAvailability;
            ::VariantInit(&vtAvailability);
            vtAvailability.vt = VT_I4;
            vtAvailability.lVal = availability;
            hr = ::SafeArrayPutElement(sa, &availIndex, &vtAvailability);
        }
    }

    if (SUCCEEDED(hr))
    {
        pvPresenceProperties->vt = VT_VARIANT | VT_ARRAY;
        pvPresenceProperties->parray = sa;
    }
    else
        ::SafeArrayDestroy(sa);

    return hr;
}

STDMETHODIMP MessengerContact::get_Property(MCONTACTPROPERTY ePropType, VARIANT *pvPropVal)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::get_ServiceId(BSTR *pbstrServiceID)
{
    return _messenger->get_MyServiceId(pbstrServiceID);
}

STDMETHODIMP MessengerContact::get_ServiceName(BSTR *pbstrServiceName)
{
    return _messenger->get_MyServiceName(pbstrServiceName);
}

STDMETHODIMP MessengerContact::get_SigninName(BSTR *pbstrSigninName)
{
    HRESULT hr;

    if (pbstrSigninName)
    {
        if (_signinName)
        {
            hr
                = ((*pbstrSigninName = ::SysAllocString(_signinName)))
                    ? S_OK
                    : E_OUTOFMEMORY;
        }
        else
        {
            *pbstrSigninName = NULL;
            hr = E_FAIL;
        }
    }
    else
        hr = RPC_X_NULL_REF_POINTER;
    return hr;
}

STDMETHODIMP MessengerContact::get_Status(MISTATUS *pMstate)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::put_Blocked(VARIANT_BOOL pBoolBlock)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::put_IsTagged(VARIANT_BOOL pBoolIsTagged)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::put_PresenceProperties(VARIANT vPresenceProperties)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::put_Property(MCONTACTPROPERTY ePropType, VARIANT vPropVal)
    STDMETHODIMP_E_NOTIMPL_STUB

STDMETHODIMP MessengerContact::QueryInterface(REFIID iid, PVOID *obj)
{
    HRESULT hr;

    if (obj)
    {
        if (IID_IMessengerContact == iid)
        {
            AddRef();
            *obj = static_cast<IMessengerContact *>(this);
            hr = S_OK;
        }
        else
            hr = DispatchImpl::QueryInterface(iid, obj);
    }
    else
        hr = E_POINTER;
    return hr;
}

BOOL MessengerContact::signinNameEquals(LPDISPATCH contact, BSTR signinName)
{
    IMessengerContact *iMessengerContact;
    HRESULT hr
        = contact->QueryInterface(
                IID_IMessengerContact,
                (PVOID *) &iMessengerContact);
    BOOL b;

    if (SUCCEEDED(hr))
    {
        BSTR contactSigninName;

        hr = iMessengerContact->get_SigninName(&contactSigninName);
        iMessengerContact->Release();
        if (SUCCEEDED(hr))
        {
            b
                = (VARCMP_EQ
                    == ::VarBstrCmp(contactSigninName, signinName, 0, 0));
            ::SysFreeString(contactSigninName);
        }
        else
            b = FALSE;
    }
    else
        b = FALSE;
    return b;
}

HRESULT MessengerContact::start()
{
	return S_OK;
}

HRESULT MessengerContact::stop()
{
    return S_OK;
}
