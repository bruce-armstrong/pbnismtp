////////////////////////////////// Includes ///////////////////////////////////

#include "stdafx.h"
#include "PJNSMTPApp.h"
#include "PJNSMTPAppDlg.h"


///////////////////////////////// Macros / Defines ////////////////////////////

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


//////////////////////////////// Implementation ///////////////////////////////

CPJNSMTPApp theApp;

BEGIN_MESSAGE_MAP(CPJNSMTPApp, CWinApp)
END_MESSAGE_MAP()

CPJNSMTPApp::CPJNSMTPApp()
{
}

BOOL CPJNSMTPApp::InitInstance()
{
  //Initialize Sockets
  if (!AfxSocketInit()) 
  {
    TRACE(_T("Failed to initialise the Winsock stack\n"));
    return FALSE;
  }

  //Settings will be stored in the registry
  SetRegistryKey(_T("PJ Naughter"));

  //Bring up the main dialog
  CPJNSMTPAppDlg dlg;
  m_pMainWnd = &dlg;
  dlg.DoModal();

  return FALSE;
}
