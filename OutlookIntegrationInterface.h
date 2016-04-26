#pragma once

#include <string>
#include <windows.h>
#include "Callbacks.h"

class OutlookIntegrationInterface
{
public:
	static OutlookIntegrationInterface* getInstance();
	static void destroy(OutlookIntegrationInterface* instance);

	void setStartConversationCallback(t_startConversation callback);

	// Callbacks ment for internal use
	STDMETHODIMP callStartConversation(std::wstring number);

private:
	OutlookIntegrationInterface();
	~OutlookIntegrationInterface();

	t_startConversation _startConversationCallback;
	
	static OutlookIntegrationInterface* _instance;
};

