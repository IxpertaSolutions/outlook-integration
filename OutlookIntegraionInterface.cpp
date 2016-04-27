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

void OutlookIntegrationInterface::setStartConversationCallback(t_startConversation callback)
{
	_startConversationCallback = callback;
}

STDMETHODIMP OutlookIntegrationInterface::callStartConversation(const wchar_t *strNumber)
{
	if (_startConversationCallback != NULL)
	{
		return _startConversationCallback(strNumber);
	}
	return E_NOTIMPL;
}

OutlookIntegrationInterface::OutlookIntegrationInterface()
	: _startConversationCallback(NULL)
{
	OutOfProcessServer::start();
}


OutlookIntegrationInterface::~OutlookIntegrationInterface()
{
	_instance = NULL;
}
