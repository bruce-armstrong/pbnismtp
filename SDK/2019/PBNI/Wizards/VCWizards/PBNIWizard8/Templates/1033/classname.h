// [!output SAFE_CLASS_NAME].h : header file for PBNI class
#ifndef [!output UPPER_CLASS_NAME]_H
#define [!output UPPER_CLASS_NAME]_H

#include <pbext.h>
[!if GEN_INTERFACE]
#include "I[!output SAFE_CLASS_NAME].h"
[!endif]

[!if VISUAL_FLAG]
class [!output SAFE_CLASS_NAME] : public [!if GEN_INTERFACE]I[!output SAFE_CLASS_NAME][!else]IPBX_VisualObject[!endif]
[!else]
class [!output SAFE_CLASS_NAME] : public [!if GEN_INTERFACE]I[!output SAFE_CLASS_NAME][!else]IPBX_NonVisualObject[!endif]
[!endif]

{
public:
	// construction/destruction
	[!output SAFE_CLASS_NAME]();
	[!output SAFE_CLASS_NAME]( IPB_Session * pSession );
	virtual ~[!output SAFE_CLASS_NAME]();

[!if GEN_INTERFACE]
	// I[!output SAFE_CLASS_NAME] interface methods
	virtual long SampleInterfaceFunction ();
[!endif]
	

	// IPBX_UserObject methods
	PBXRESULT Invoke
	(
		IPB_Session * session,
		pbobject obj,
		pbmethodID mid,
		PBCallInfo * ci
	);

   void Destroy();

[!if VISUAL_FLAG]
   // IPBX_VisualObject methods
   LPCTSTR GetWindowClassName();

   HWND CreateControl
	(
		DWORD dwExStyle,      // extended window style
		LPCTSTR lpWindowName, // window name
		DWORD dwStyle,        // window style
		int x,                // horizontal position of window
		int y,                // vertical position of window
		int nWidth,           // window width
		int nHeight,          // window height
		HWND hWndParent,      // handle to parent or owner window
		HINSTANCE hInstance   // handle to application instance
	);

   static void RegisterClass(); //Register the Window class
   static void UnregisterClass();//UnRegister the Window class
   static LRESULT CALLBACK WindowProc(HWND hwnd,UINT uMsg,WPARAM wParam,LPARAM lParam);//Callback Window Procedure
   static void SetDLLHandle(HANDLE hndl);//Set the dll HINSTANCE

[!endif]

	// PowerBuilder method wrappers
	enum Function_Entrys
	{
		mid_Hello = 0,
		// TODO: add enum entries for each callable method
		NO_MORE_METHODS
	};

[!if GLOBAL]
 	// global methods callable from PowerBuilder
	PBXRESULT HelloGlobal( PBCallInfo * ci );
[!endif]

protected:
 	// methods callable from PowerBuilder
	PBXRESULT Hello( PBCallInfo * ci );

protected:
    // member variables
    IPB_Session * m_pSession;
 [!if VISUAL_FLAG]
	static HINSTANCE m_Handle;
	HWND		d_hwnd;
 [!endif]
};

#endif	// !defined([!output UPPER_CLASS_NAME]_H)