#include "stdafx.h"
#include "PJNSMTPApp.h"
#include "SendingDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CSendingDlg::CSendingDlg(CWnd* pParent)	: CDialog(CSendingDlg::IDD, pParent),
                                          m_bSafeToClose(FALSE),
                                          m_bAbort(FALSE)
{
}

void CSendingDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);

  DDX_Control(pDX, IDC_PROGRESS1, m_ctrlProgress);
}

BEGIN_MESSAGE_MAP(CSendingDlg, CDialog)
END_MESSAGE_MAP()

BOOL CSendingDlg::OnInitDialog() 
{
  //Let the base class do its thing
  CDialog::OnInitDialog();

  //Set the initialized event	
  m_evtInitialized.SetEvent();  
  
  return TRUE;
}

void CSendingDlg::OnCancel() 
{
  if (m_bSafeToClose)
    CDialog::OnCancel();
  else
  {
    //Just set the abort flag to TRUE
    m_bAbort = TRUE;	
  }
}
