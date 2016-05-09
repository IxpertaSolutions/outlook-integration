#pragma once

#include <string>
#include <windows.h>
#include "Callbacks.h"

class OutlookIntegrationInterface
{
public:
	static OutlookIntegrationInterface* getInstance();

	void destroy();
	void setStartConversationCallback(startConversation_t callback);

	// Callbacks ment for internal use
	int callStartConversation(const wchar_t *strNumber);

	// Set callback for logging purposes
	void setLoggingFunc(log_t);
	void log(const wchar_t *, const int, const wchar_t *, const wchar_t *);

private:
	OutlookIntegrationInterface();
	~OutlookIntegrationInterface();

	startConversation_t _startConversationCallback;
	log_t _logFunction;
	
	static OutlookIntegrationInterface* _instance;
};

