#include "OutlookIntegrationExported.h"
#include "OutlookIntegrationInterface.h"

extern "C" {

DLL_EXPORT void* outlook_integration_interface_get(log_t logFunction)
{
	OutlookIntegrationInterface * instance;
	instance = OutlookIntegrationInterface::getInstance();
	instance->setLoggingFunc(logFunction);
	return instance;
}

DLL_EXPORT void outlook_integration_interface_set_conversation_start_callback(void *p, startConversation_t callback)
{
	OutlookIntegrationInterface *_interface = (OutlookIntegrationInterface*)p;
	_interface->setStartConversationCallback(callback);
}

DLL_EXPORT void outlook_integration_interface_destroy(void *p)
{
	OutlookIntegrationInterface *_interface = (OutlookIntegrationInterface*)p;
	_interface->destroy();
}

}