#pragma once

#ifndef STRICT
#define STRICT
#endif

#ifndef WINVER
#define WINVER 0x0600
#endif

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0600
#endif

#ifndef _WIN32_WINDOWS
#define _WIN32_WINDOWS 0x0600
#endif

#ifndef _WIN32_IE
#define _WIN32_IE 0x0600
#endif

#define VC_EXTRALEAN //Exclude rarely-used stuff from Windows headers

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS //some CString constructors will be explicit

#define _AFX_ALL_WARNINGS // turns off MFC's hiding of some common and often safely ignored warning messages

#ifndef _SECURE_ATL
#define _SECURE_ATL 1 //Use the Secure C Runtime in ATL
#endif

//#define CPJNSMTP_NOSSL //Uncomment this line to try out the SMTP classes without SSL support
//#define CPJNSMTP_NONTLM //Uncomment this line to try out the SMTP classes without NTLM support

//#define CPJNSMTP_MFC_EXTENSIONS //Uncomment this line if you want to use MFC

#ifndef CPJNSMTP_MFC_EXTENSIONS
  #include <atlbase.h>
  #include <atlstr.h>
  #include <atlsocket.h>
#else
  #define CWSOCKET_MFC_EXTENSIONS
  #define CPJNMD5_MFC_EXTENSIONS
  #define CNTLMCLIENTAUTH_MFC_EXTENSIONS
  #include <afxwin.h>
  #include <afxtempl.h>
  #include <afxsock.h>
  #include <afxmt.h>
  #include <afxcmn.h>
  #include <afxdlgs.h>
#endif //#ifndef CPJNSMTP_MFC_EXTENSIONS

#include <locale.h>
#include <wincrypt.h>
#include <atlenc.h>
#include <atlfile.h>
#include <wininet.h>
#include <WinDNS.h>
#include <WS2tcpip.h>

#ifndef CPJNSMTP_NOSSL
  #include <schannel.h>
  #include <cryptuiapi.h>
  #include <vector>
  #include <string>
  #define SECURITY_WIN32
  #include <sspi.h>
#endif //#ifndef CPJNSMTP_NOSSL

#ifndef CPJNSMTP_NONTLM
  #define SECURITY_WIN32
  #include <sspi.h>
#endif //#ifndef CPJNSMTP_NONTLM

//We need Crypt32 for DPAPI support in the sample app
#pragma comment(lib, "crypt32.lib")

#ifndef CPJNSMTP_MFC_EXTENSIONS
  #include <exception>
  #include <string>
  #include <vector>
  #include <sstream>
  #include <algorithm>
#endif //#ifndef CPJNSMTP_MFC_EXTENSIONS

//Pull in support for Commctrl v6
#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif

