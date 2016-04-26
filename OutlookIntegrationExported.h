#pragma once

#include "Callbacks.h"

#ifdef MSOFFICECOMM_EXPORTS
#define DLL_EXPORT __declspec(dllexport)
#else
#define DLL_EXPORT __declspec(dllimport)
#endif

extern "C" {

DLL_EXPORT void* outlook_integration_interface_get();

DLL_EXPORT void outlook_integration_interface_set_conversation_start_callback(void *, t_startConversation);

DLL_EXPORT void outlook_integration_interface_destroy(void *);

}