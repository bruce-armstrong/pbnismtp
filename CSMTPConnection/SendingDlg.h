#pragma once

class CSendingDlg : public CDialog
{
public:
//Constructors / Destructors
  CSendingDlg(CWnd* pParent = NULL);   // standard constructor

//Methods
  void SetSafeToClose() { m_bSafeToClose = TRUE; };
  BOOL AbortFlag() const { return m_bAbort; };

//Member variables
  CEvent m_evtInitialized; //set when we have been constructed and have a HWND

  enum { IDD = IDD_PROGRESS };
  CProgressCtrl	m_ctrlProgress;

protected:
  virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

  virtual BOOL OnInitDialog();
  virtual void OnCancel();

  DECLARE_MESSAGE_MAP()

  BOOL m_bSafeToClose;
  BOOL m_bAbort;
};
