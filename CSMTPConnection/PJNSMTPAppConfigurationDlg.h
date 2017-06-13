#pragma once

class CPJNSMTPAppConfigurationDlg : public CDialog
{
public:
//Constructors / Destructors
  CPJNSMTPAppConfigurationDlg(CWnd* pParent = NULL);   // standard constructor

//Member variables
  enum { IDD = IDD_CONFIGURATION };
  CComboBox	m_ctrlEncoding;
  CComboBox	m_ctrlIPAddresses;
  CComboBox	m_ctrlAuthenticate;
  CStatic	m_ctrlPromptUsername;
  CStatic	m_ctrlPromptPassword;
  CEdit	m_ctrlUsername;
  CEdit	m_ctrlPassword;
  CString	m_sAddress;
  CString	m_sHost;
  CString	m_sName;
  int		m_nPort;
  CString	m_sPassword;
  CString	m_sUsername;
  BOOL	m_bAutoDial;
  CString	m_sBoundIP;
  CString m_sEncodingFriendly;
  BOOL	m_bMime;
  BOOL	m_bHTML;
  DWORD m_ConnectionType;
  DWORD m_SSLProtocol;
  BOOL	m_bDSN;
  BOOL	m_bDSNDelay;
  BOOL	m_bDSNFailure;
  BOOL	m_bDSNSuccess;
  BOOL	m_bDSNHeaders;
  DWORD	m_Auth;
  CString m_sEncodingCharset;
  DWORD m_Priority;
  BOOL  m_bMDN;
  BOOL  m_bInitialClientResponse;

protected:
  virtual void DoDataExchange(CDataExchange* pDX);
  virtual BOOL OnInitDialog();

  int CBAddStringAndData(CWnd* pDlg, int nIDC, LPCTSTR pszString, DWORD_PTR dwItemData);
  BOOL AddLocalIpsToBindCombo();
  static void DDX_CBData(CDataExchange* pDX, int nIDC, DWORD& dwItemData);

  afx_msg void OnSelchangeAuthenticate();
  afx_msg void OnConnectionType();
  afx_msg void OnDSN();

  DECLARE_MESSAGE_MAP()
};

