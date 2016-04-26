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

void OutlookIntegrationInterface::destroy(OutlookIntegrationInterface * instance)
{
	delete _instance;
}

void OutlookIntegrationInterface::setStartConversationCallback(t_startConversation callback)
{
	_startConversationCallback = callback;
}

STDMETHODIMP OutlookIntegrationInterface::callStartConversation(std::wstring number)
{
	if (_startConversationCallback != NULL)
	{
		return _startConversationCallback(number);
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
