#pragma once

#include "Callbacks.h"

#ifdef MSOFFICECOMM_EXPORTS
#define DLL_EXPORT __declspec(dllexport)
#else
#define DLL_EXPORT __declspec(dllimport)
#endif

extern "C" {

DLL_EXPORT void* outlook_integration_interface_get(log_t);

DLL_EXPORT void outlook_integration_interface_set_conversation_start_callback(void *, startConversation_t);

DLL_EXPORT void outlook_integration_interface_destroy(void *);

}
