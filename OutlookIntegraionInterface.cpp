#include "OutlookIntegrationInterface.h"

#include "OutOfProcessServer.h"

OutlookIntegrationInterface* OutlookIntegrationInterface::_instance = NULL;

OutlookIntegrationInterface * OutlookIntegrationInterface::getInstance()
{
	if (_instance == NULL)
	{
		_instance = new OutlookIntegrationInterface;
	}
	return _instance;
}

void OutlookIntegrationInterface::destroy()
{
	delete _instance;
	_instance = NULL;
}

void OutlookIntegrationInterface::setStartConversationCallback(startConversation_t callback)
{
	_startConversationCallback = callback;
}

int OutlookIntegrationInterface::callStartConversation(const wchar_t *strNumber)
{
	if (_startConversationCallback != NULL)
	{
		return _startConversationCallback(strNumber);
	}
	return 5; // Is transalted to E_NOTIMPL
}

void OutlookIntegrationInterface::setLoggingFunc(log_t f)
{
	_logFunction = f;
}

void OutlookIntegrationInterface::log(const wchar_t *file,
	const int line,
	const wchar_t *function_name,
	const wchar_t *msg)
{
	if (_logFunction != NULL)
		_logFunction(file, line, function_name, msg);
}

OutlookIntegrationInterface::OutlookIntegrationInterface()
	: _startConversationCallback(NULL),
	_logFunction(NULL)
{
	OutOfProcessServer::start();
}


OutlookIntegrationInterface::~OutlookIntegrationInterface()
{
	_instance = NULL;
}
