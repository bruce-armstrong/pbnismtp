// [!output SAFE_CLASS_NAME].cpp : PBNI class
#include "[!output SAFE_CLASS_NAME].h"

[!if VISUAL_FLAG]
HINSTANCE [!output SAFE_CLASS_NAME]::m_Handle = 0;
[!endif]

// default constructor
[!output SAFE_CLASS_NAME]::[!output SAFE_CLASS_NAME]()
{
}

[!output SAFE_CLASS_NAME]::[!output SAFE_CLASS_NAME]( IPB_Session * pSession )
:m_pSession( pSession )
{
}

// destructor
[!output SAFE_CLASS_NAME]::~[!output SAFE_CLASS_NAME]()
{
}

// method called by PowerBuilder to invoke PBNI class methods
PBXRESULT [!output SAFE_CLASS_NAME]::Invoke
(
	IPB_Session * session,
	pbobject obj,
	pbmethodID mid,
	PBCallInfo * ci
)
{
   PBXRESULT pbxr = PBX_OK;

	switch ( mid )
	{
		case mid_Hello:
			pbxr = this->Hello( ci );
         break;
		
		// TODO: add handlers for other callable methods

		default:
			pbxr = PBX_E_INVOKE_METHOD_AMBIGUOUS;
	}

	return pbxr;
}

[!if VISUAL_FLAG]
// IPBX_VisualObject method
LPCTSTR [!output SAFE_CLASS_NAME]::GetWindowClassName()
{
   return LPCTSTR( "[!output SAFE_CLASS_NAME]" );
}


void [!output SAFE_CLASS_NAME]::SetDLLHandle(HANDLE hndl)
{
    m_Handle = (HINSTANCE)hndl;	
}

void [!output SAFE_CLASS_NAME]::RegisterClass()
{
   WNDCLASS wndclass;

   wndclass.style = CS_GLOBALCLASS | CS_DBLCLKS;
   wndclass.lpfnWndProc = (WNDPROC)[!output SAFE_CLASS_NAME]::WindowProc;
   wndclass.cbClsExtra = 0;
   wndclass.cbWndExtra = 0;
   wndclass.hInstance = m_Handle;
   wndclass.hIcon = NULL;
   wndclass.hCursor = LoadCursor (NULL, IDC_ARROW);
   wndclass.hbrBackground =(HBRUSH) (COLOR_WINDOW + 1);
   wndclass.lpszMenuName = NULL;
   wndclass.lpszClassName = _T("[!output SAFE_CLASS_NAME]");

   ::RegisterClass (&wndclass);
}


void [!output SAFE_CLASS_NAME]::UnregisterClass()
{
	::UnregisterClass(_T("[!output SAFE_CLASS_NAME]"),m_Handle );
}

LRESULT CALLBACK [!output SAFE_CLASS_NAME]::WindowProc(
                                   HWND hwnd,
                                   UINT uMsg,
                                   WPARAM wParam,
                                   LPARAM lParam
                                   )
{
     
   switch(uMsg) {

   case WM_CREATE:
      return 0;

   case WM_SIZE:
      return 0;

   case WM_COMMAND:
      return 0;

   case WM_PAINT: {
      PAINTSTRUCT ps;
      HDC hdc = BeginPaint(hwnd, &ps);
      RECT rc;
      GetClientRect(hwnd, &rc);
 
      Rectangle(hdc, rc.left, rc.top, rc.right-rc.left,
                rc.bottom-rc.top);

	   DrawText(hdc, _T("Hello World"),11, &rc, 
               DT_CENTER|DT_VCENTER|DT_SINGLELINE);
      EndPaint(hwnd, &ps);
      }
      return 0;

   }
   return DefWindowProc(hwnd, uMsg, wParam, lParam);
}


// IPBX_VisualObject method
HWND [!output SAFE_CLASS_NAME]::CreateControl
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
)
{
   // call ::CreateWindow, ::CreateWindowEx or the appropriate creation method supplied by a control class
   // e.g.:
   d_hwnd = CreateWindowEx(
      dwExStyle,
      _T("[!output SAFE_CLASS_NAME]"),
      lpWindowName,
      dwStyle,
		x,
      y,
      nWidth,
      nHeight,
      hWndParent,
      NULL,
      hInstance,
      NULL
   );   
    
   return d_hwnd;
}
[!endif]

void [!output SAFE_CLASS_NAME]::Destroy() 
{
   delete this;
}

// Method callable from PowerBuilder
PBXRESULT [!output SAFE_CLASS_NAME]::Hello( PBCallInfo * ci )
{
	PBXRESULT	pbxr = PBX_OK;

	// return value
	ci->returnValue->SetString( _T("Hello") );

	[!if VISUAL_FLAG]
	RECT rc;
	GetClientRect(d_hwnd, &rc);
	InvalidateRect(d_hwnd, &rc, TRUE);
	[!endif]

	return pbxr;
}

[!if GLOBAL]
// Method callable from PowerBuilder
PBXRESULT [!output SAFE_CLASS_NAME]::HelloGlobal( PBCallInfo * ci )
{
	PBXRESULT	pbxr = PBX_OK;

	// return value
	ci->returnValue->SetString( _T("Hello Global") );

	return pbxr;
}
[!endif]

[!if GEN_INTERFACE]
// I[!output SAFE_CLASS_NAME]
long [!output SAFE_CLASS_NAME]::SampleInterfaceFunction()
{
  long lRet = 0;

  return lRet;
}
[!endif]
