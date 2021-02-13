// [!output PROJECT_NAME].cpp : defines the entry point for the PBNI DLL.
//Global: [!output GLOBAL]
//Unicode: [!output PB10UNICODE]
#include "[!output SAFE_PROJECT_NAME].h"
#include "[!output SAFE_CLASS_NAME].h"

[!if GLOBAL]
// global variables
[!output SAFE_CLASS_NAME] * g_pGlobal = NULL;    // used by global PB functions
[!endif]

BOOL APIENTRY DllMain
(
   HANDLE hModule,
   DWORD ul_reason_for_all,
   LPVOID lpReserved
)
{	[!if VISUAL_FLAG]

	switch(ul_reason_for_all)
	{
	case DLL_PROCESS_ATTACH:
		[!output SAFE_CLASS_NAME]::RegisterClass();
		[!output SAFE_CLASS_NAME]::SetDLLHandle(hModule);
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		[!output SAFE_CLASS_NAME]::UnregisterClass();
		break;
	}
	[!endif]

	[!if NON_VISUAL_FLAG]
	
	switch(ul_reason_for_all)
	{	
	case DLL_PROCESS_ATTACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:		
	case DLL_PROCESS_DETACH:
		break;
	}
	[!endif]
		
	return TRUE;
}

// The description returned from PBX_GetDescription is used by
// the PBX2PBD tool to create pb groups for the PBNI class.
//
// + The description can contain more than one class definition.
// + A class definition can reference any class definition that
//   appears before it in the description.
// + The PBNI class must inherit from a class that inherits from
//   NonVisualObject, such as Transaction or Exception

PBXEXPORT LPCTSTR PBXCALL PBX_GetDescription()
{
   // combined class description
   static const TCHAR classDesc[] = {
      /* [!output SAFE_CLASS_NAME] */
      _T("class [!if VISUAL_FLAG]u[!else]n[!endif]_cpp_[!output LOWER_CASE_PROJECT_NAME] from [!if VISUAL_FLAG]user[!else]nonvisual[!endif]object\n") \
      _T("   function string of_hello()\n") \

      // TODO: add method declarations for this class

      _T("end class\n")

      // TODO: additional class and method declarations

      [!if GLOBAL]
      // Global functions
      _T("globalfunctions\n") \
      _T("   function string f_hello()\n") \
      _T("end globalfunctions\n") \
      [!endif]
   };

   return (LPCTSTR)classDesc;
}

[!if VISUAL_FLAG]
// PBX_CreateVisualObject is called by PowerBuilder to create a C++ class
// that corresponds to the PBNI visual user object created by PowerBuilder.
PBXEXPORT PBXRESULT PBXCALL PBX_CreateVisualObject
(
   IPB_Session * session,
   pbobject obj,
   LPCTSTR className,
   IPBX_VisualObject ** nvobj
)
{
   // The name must not contain upper case
   if (_tcscmp(className, _T("u_cpp_[!output LOWER_CASE_PROJECT_NAME]")) == 0)
      *nvobj = new [!output SAFE_CLASS_NAME](session);
   return PBX_OK;
}
[!else]
// PBX_CreateNonVisualObject is called by PowerBuilder to create a C++ class
// that corresponds to the PBNI class created by PowerBuilder.
PBXEXPORT PBXRESULT PBXCALL PBX_CreateNonVisualObject
(
   IPB_Session * session,
   pbobject obj,
   LPCTSTR className,
   IPBX_NonVisualObject ** nvobj
)
{
   // The name must not contain upper case
   if (_tcscmp(className, _T("n_cpp_[!output LOWER_CASE_PROJECT_NAME]")) == 0)
      *nvobj = new [!output SAFE_CLASS_NAME](session);

   return PBX_OK;
}
[!endif]

[!if GLOBAL]
PBXEXPORT PBXRESULT PBXCALL PBX_InvokeGlobalFunction
(
   IPB_Session * session,
   LPCTSTR functionName,
   PBCallInfo * ci
)
{
   PBXRESULT pbxr = PBX_OK;

   // first time invoked -- instantiate class
   if (g_pGlobal == NULL) g_pGlobal = new [!output SAFE_CLASS_NAME](session);

   // The name must not contain upper case
   if (_tcscmp(functionName, _T("f_hello")) == 0)
      pbxr = g_pGlobal->HelloGlobal(ci);
   else
      pbxr = PBX_E_INVOKE_METHOD_AMBIGUOUS;

   return pbxr;
}
[!endif]