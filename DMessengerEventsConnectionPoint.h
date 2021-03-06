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
#ifndef _JMSOFFICECOMM_DMESSENGEREVENTSCONNECTIONPOINT_H_
#define _JMSOFFICECOMM_DMESSENGEREVENTSCONNECTIONPOINT_H_

#include "ConnectionPoint.h"
#include <msgrua.h>

class DMessengerEventsConnectionPoint
    : public ConnectionPoint<DMessengerEvents, DIID_DMessengerEvents>
{
public:
    DMessengerEventsConnectionPoint(IConnectionPointContainer *container)
        : ConnectionPoint(container) {}
    virtual ~DMessengerEventsConnectionPoint() {}

    STDMETHODIMP OnContactStatusChange(LPDISPATCH pMContact, MISTATUS mStatus);
};

#endif /* #ifndef _JMSOFFICECOMM_DMESSENGEREVENTSCONNECTIONPOINT_H_ */
