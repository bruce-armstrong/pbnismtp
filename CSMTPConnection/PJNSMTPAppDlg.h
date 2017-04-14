#pragma once

#include "PJNSMTP.h"

class CPJNSMTPAppDlg : public CDialog
{
public:
  CPJNSMTPAppDlg(CWnd* pParent = NULL);

  enum { IDD = IDD_MAIN_DIALOG };
  CEdit	                                   m_wndBody;
  CString	                                 m_sBCC;
  CString	                                 m_sBody;
  CString	                                 m_sCC;
  CString	                                 m_sFile;
  CString	                                 m_sSubject;
  CString	                                 m_sTo;
  BOOL	                                   m_bDirectly;
  BOOL	                                   m_bDNSLookup;
  CString	                                 m_sAddress;
  CString	                                 m_sHost;
  CString	                                 m_sName;
  int                                      m_nPort;
  CPJNSMTPConnection::AuthenticationMethod m_Auth;
  CString                                  m_sUsername;
  CString                                  m_sPassword;
  BOOL                                     m_bAutoDial;
  CString                                  m_sBoundIP;
  CString                                  m_sEncodingFriendly;
  CString                                  m_sEncodingCharset;
  BOOL                                     m_bMime;
  BOOL                                     m_bHTML;
  CPJNSMTPMessage::PRIORITY                m_Priority;
  CPJNSMTPConnection::ConnectionType       m_ConnectionType;
  CPJNSMTPConnection::SSLProtocol          m_SSLProtocol;
  CPJNSMTPConnection                       m_Connection;
  BOOL                                     m_bDSN;
  BOOL                                     m_bDSNSuccess;
  BOOL                                     m_bDSNFailure;
  BOOL                                     m_bDSNDelay;
  BOOL                                     m_bDSNHeaders;
  BOOL                                     m_bMDN;

  virtual void DoDataExchange(CDataExchange* pDX);

protected:
  HICON m_hIcon;

  static UINT _SendingWorkerThread(LPVOID lpParam);
  CPJNSMTPMessage* CreateMessage();

  virtual BOOL OnInitDialog();
  afx_msg void OnPaint();
  afx_msg HCURSOR OnQueryDragIcon();
  afx_msg void OnConfiguration();
  afx_msg void OnSend();
  afx_msg void OnBrowseFile();
  afx_msg void OnDestroy();
  afx_msg void OnSendFromDisk();
  afx_msg void OnSave();

  DECLARE_MESSAGE_MAP()
};

