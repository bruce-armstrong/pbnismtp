#pragma once

#ifndef __AFXWIN_H__
  #error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

class CPJNSMTPApp : public CWinApp
{
public:
//Constructors / Destructors
  CPJNSMTPApp();

protected:
  virtual BOOL InitInstance();
  
  DECLARE_MESSAGE_MAP()
};
