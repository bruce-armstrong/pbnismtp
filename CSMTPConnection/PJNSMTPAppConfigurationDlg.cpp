#include "stdafx.h"
#include "PJNSMTP.h"
#include "PJNSMTPApp.h"
#include "PJNSMTPAppConfigurationDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

struct _MC_charset
{
  TCHAR* pszFriendlyName;
  TCHAR* pszCharset;
};

_MC_charset g_charset[] =
{
  { _T("Arabic (Windows)"), _T("windows-1256") },
  { _T("Baltic (Windows)"), _T("windows-1257") },
  { _T("Central European (ISO)"), _T("iso-8859-2") },
  { _T("Central European (Windows)"), _T("windows-1250") },
  { _T("Chinese Simplified (GB2312)"), _T("gb2312") },
  { _T("Chinese Simplified (HZ)"), _T("hz-gb-2312") },
  { _T("Chinese Traditional (Big5)"), _T("big5") },
  { _T("Cyrilic (KOI8-R)"), _T("koi8-r") },
  { _T("Cyrillic (Windows)"), _T("windows-1251") },
  { _T("Greek (Windows)"), _T("windows-1253") },
  { _T("Hebrew (Windows)"), _T("windows-1255") },
  { _T("Japanese (JIS)"), _T("iso-2022-jp") },
  { _T("Korean"), _T("ks_c_5601") },
  { _T("Korean (EUC)"), _T("euc-kr") },
  { _T("Latin 9 (ISO)"), _T("iso-8859-15") },
  { _T("Thai (Windows)"), _T("windows-874") },
  { _T("Turkish (Windows)"), _T("windows-1254") },
  { _T("Unicode (UTF-7)"), _T("utf-7") },
  { _T("Unicode (UTF-8)"), _T("utf-8") },
  { _T("Vietnamese (Windows)"), _T("windows-1258") },
  { _T("Turkish (Windows)"), _T("windows-1254") },
  { _T("Western European (ISO)"), _T("iso-8859-1") },
  { _T("Western European (Windows)"), _T("windows-1252") },
};

BEGIN_MESSAGE_MAP(CPJNSMTPAppConfigurationDlg, CDialog)
  ON_CBN_SELCHANGE(IDC_AUTH, &CPJNSMTPAppConfigurationDlg::OnSelchangeAuthenticate)
  ON_CBN_SELCHANGE(IDC_CONNECTIONTYPE, &CPJNSMTPAppConfigurationDlg::OnConnectionType)
  ON_BN_CLICKED(IDC_DSN, &CPJNSMTPAppConfigurationDlg::OnDSN)
END_MESSAGE_MAP()

CPJNSMTPAppConfigurationDlg::CPJNSMTPAppConfigurationDlg(CWnd* pParent)	: CDialog(CPJNSMTPAppConfigurationDlg::IDD, pParent),
                                                                          m_bAutoDial(FALSE),
                                                                          m_bMime(FALSE),
                                                                          m_bHTML(FALSE),
                                                                          m_ConnectionType(CPJNSMTPConnection::PlainText),
                                                                          m_SSLProtocol(CPJNSMTPConnection::OSDefault),
                                                                          m_bDSN(FALSE),
                                                                          m_bDSNDelay(FALSE),
                                                                          m_bDSNFailure(FALSE),
                                                                          m_bDSNSuccess(FALSE),
                                                                          m_bDSNHeaders(TRUE),
                                                                          m_nPort(25), //Use the default SMTP port
                                                                          m_Auth(CPJNSMTPConnection::AUTH_NONE),
                                                                          m_sEncodingFriendly(_T("Western European (ISO)")),
                                                                          m_sEncodingCharset(_T("iso-8859-1")),
                                                                          m_Priority(CPJNSMTPMessage::NoPriority),
                                                                          m_bMDN(FALSE),
                                                                          m_bInitialClientResponse(FALSE)
{
}

int CPJNSMTPAppConfigurationDlg::CBAddStringAndData(CWnd* pDlg, int nIDC, LPCTSTR pszString, DWORD_PTR dwItemData)
{
  //Validate our parameters
  AFXASSUME(pDlg);

  CComboBox* pComboBox = static_cast<CComboBox*>(pDlg->GetDlgItem(nIDC));
  AFXASSUME(pComboBox);

  int nInserted = pComboBox->AddString(pszString);
  if (nInserted != CB_ERR)
    nInserted = pComboBox->SetItemData(nInserted, dwItemData);

  return nInserted;
}

void CPJNSMTPAppConfigurationDlg::DoDataExchange(CDataExchange* pDX)
{
  //Let the base class do its thing
  CDialog::DoDataExchange(pDX);

  if (!pDX->m_bSaveAndValidate)
  {
    //Add the login methods to the combo box
    CBAddStringAndData(this, IDC_AUTH, _T("None"), CPJNSMTPConnection::AUTH_NONE);
    CBAddStringAndData(this, IDC_AUTH, _T("CRAM MD5"), CPJNSMTPConnection::AUTH_CRAM_MD5);
    CBAddStringAndData(this, IDC_AUTH, _T("AUTH LOGIN"), CPJNSMTPConnection::AUTH_LOGIN);
    CBAddStringAndData(this, IDC_AUTH, _T("AUTH PLAIN"), CPJNSMTPConnection::AUTH_PLAIN);
  #ifndef CPJNSMTP_NONTLM
    CBAddStringAndData(this, IDC_AUTH, _T("NTLM"), CPJNSMTPConnection::AUTH_NTLM);
  #endif //#ifndef CPJNSMTP_NONTLM
    CBAddStringAndData(this, IDC_AUTH, _T("XOAUTH2"), CPJNSMTPConnection::AUTH_XOAUTH2);
    CBAddStringAndData(this, IDC_AUTH, _T("Auto Detect"), CPJNSMTPConnection::AUTH_AUTO);

    //Add the charset methods to the combo box
    for (int i=0; i<sizeof(g_charset)/sizeof(_MC_charset); i++)
      CBAddStringAndData(this, IDC_ENCODING, g_charset[i].pszFriendlyName, i);

    //Add the priority values to the combo box
    CBAddStringAndData(this, IDC_PRIORITY, _T("None Defined"), CPJNSMTPMessage::NoPriority);
    CBAddStringAndData(this, IDC_PRIORITY, _T("Low"), CPJNSMTPMessage::LowPriority);
    CBAddStringAndData(this, IDC_PRIORITY, _T("Normal"), CPJNSMTPMessage::NormalPriority);
    CBAddStringAndData(this, IDC_PRIORITY, _T("High"), CPJNSMTPMessage::HighPriority);

     //Add the connection type values to the combo box 
    CBAddStringAndData(this, IDC_CONNECTIONTYPE, _T("Plain Text"), CPJNSMTPConnection::PlainText);
  #ifndef CPJNSMTP_NOSSL
    CBAddStringAndData(this, IDC_CONNECTIONTYPE, _T("SSL / TLS"), CPJNSMTPConnection::SSL_TLS);
    CBAddStringAndData(this, IDC_CONNECTIONTYPE, _T("STARTTLS"), CPJNSMTPConnection::STARTTLS);
    CBAddStringAndData(this, IDC_CONNECTIONTYPE, _T("Auto Upgrade to STARTTLS"), CPJNSMTPConnection::AutoUpgradeToSTARTTLS);
  
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("SSL v2 or SSL v3"), CPJNSMTPConnection::SSLv2orv3);
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("SSL v2"), CPJNSMTPConnection::SSLv2);
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("SSL v3"), CPJNSMTPConnection::SSLv3);
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("TLS v1.0"), CPJNSMTPConnection::TLSv1);
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("TLS v1.1"), CPJNSMTPConnection::TLSv1_1);
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("TLS v1.2"), CPJNSMTPConnection::TLSv1_2);
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("TLS v1.0 or TLS v1.1 or TLS v1.2"), CPJNSMTPConnection::AnyTLS);
    CBAddStringAndData(this, IDC_SSL_PROTOCOL, _T("OS Default"), CPJNSMTPConnection::OSDefault);
  #endif //#ifndef CPJNSMTP_NOSSL
  
  }

  //Add the IP address to the combo
  if (!pDX->m_bSaveAndValidate)
  {
    DDX_Control(pDX, IDC_IPS, m_ctrlIPAddresses);
    AddLocalIpsToBindCombo();

    if (m_sBoundIP == _T(""))
      m_sBoundIP = _T("ANY_IP_ADDRESS");
  }

  DDX_Control(pDX, IDC_ENCODING, m_ctrlEncoding);
  DDX_Control(pDX, IDC_AUTH, m_ctrlAuthenticate);
  DDX_Control(pDX, IDC_PROMPT_USERNAME, m_ctrlPromptUsername);
  DDX_Control(pDX, IDC_PROMPT_PASSWORD, m_ctrlPromptPassword);
  DDX_Control(pDX, IDC_USERNAME, m_ctrlUsername);
  DDX_Control(pDX, IDC_PASSWORD, m_ctrlPassword);
  DDX_Text(pDX, IDC_ADDRESS, m_sAddress);
  DDX_Text(pDX, IDC_HOST, m_sHost);
  DDX_Text(pDX, IDC_NAME, m_sName);
  DDX_Text(pDX, IDC_PORT, m_nPort);
  DDV_MinMaxInt(pDX, m_nPort, 1, 32000);
  DDX_Text(pDX, IDC_PASSWORD, m_sPassword);
  DDX_Text(pDX, IDC_USERNAME, m_sUsername);
  DDX_Check(pDX, IDC_AUTOCONNECT, m_bAutoDial);
  DDX_CBString(pDX, IDC_IPS, m_sBoundIP);
  DDX_CBString(pDX, IDC_ENCODING, m_sEncodingFriendly);
  DDX_Check(pDX, IDC_MIME, m_bMime);
  DDX_Check(pDX, IDC_HTML, m_bHTML);
  DDX_CBData(pDX, IDC_CONNECTIONTYPE, m_ConnectionType);
  DDX_CBData(pDX, IDC_SSL_PROTOCOL, m_SSLProtocol);
  DDX_Check(pDX, IDC_DSN, m_bDSN);
  DDX_Check(pDX, IDC_DSN_DELAY, m_bDSNDelay);
  DDX_Check(pDX, IDC_DSN_FAILURE, m_bDSNFailure);
  DDX_Check(pDX, IDC_DSN_SUCCESS, m_bDSNSuccess);
  DDX_Check(pDX, IDC_DSN_HEADERS, m_bDSNHeaders);
  DDX_CBData(pDX, IDC_AUTH, m_Auth);
  DDX_CBData(pDX, IDC_PRIORITY, m_Priority);
  DDX_Check(pDX, IDC_MDN, m_bMDN);
  DDX_Check(pDX, IDC_INITIAL_CLIENT_RESPONSE, m_bInitialClientResponse);

  if (pDX->m_bSaveAndValidate)
  {
    int nSel = m_ctrlEncoding.GetCurSel();
    if (nSel != -1)
      m_sEncodingCharset = g_charset[nSel].pszCharset;

    if (m_sAddress.IsEmpty())
    {
      AfxMessageBox(_T("Please specify a valid email address"));
      pDX->PrepareEditCtrl(IDC_ADDRESS);
      pDX->Fail();
    }
    if (m_sHost.IsEmpty())
    {
      AfxMessageBox(_T("Please specify a valid SMTP host"));
      pDX->PrepareEditCtrl(IDC_HOST);
      pDX->Fail();
    }

    if (m_Auth != CPJNSMTPConnection::AUTH_NONE && m_Auth != CPJNSMTPConnection::AUTH_NTLM)
    {
      if (m_sUsername.IsEmpty())
      {
        AfxMessageBox(_T("Please specify a valid username"));
        pDX->PrepareEditCtrl(IDC_USERNAME);
        pDX->Fail();
      }
    }
  }

  if (!pDX->m_bSaveAndValidate)
  {
    OnSelchangeAuthenticate();
    OnDSN();
  }
  else
  {
    if (m_sBoundIP == _T("ANY_IP_ADDRESS"))
      m_sBoundIP = _T("");
  }
}

BOOL CPJNSMTPAppConfigurationDlg::OnInitDialog() 
{
  //Let the base class do its thing
  CDialog::OnInitDialog();
  
  //Force the connection type to plain text if we are not using SSL
#ifdef CPJNSMTP_NOSSL
  m_ConnectionType = CPJNSMTPConnection::PlainText;
  CDataExchange DX(this, FALSE);
  DDX_CBData(&DX, IDC_CONNECTIONTYPE, m_ConnectionType);
  GetDlgItem(IDC_CONNECTIONTYPE)->EnableWindow(FALSE);
  GetDlgItem(IDC_SSL_PROTOCOL)->EnableWindow(FALSE);
#else
  GetDlgItem(IDC_SSL_PROTOCOL)->EnableWindow(m_ConnectionType != CPJNSMTPConnection::PlainText);
#endif //#ifdef CPJNSMTP_NOSSL

  return TRUE;
}

BOOL CPJNSMTPAppConfigurationDlg::AddLocalIpsToBindCombo()
{
  //Enumerate the IP's on this machine and add them to the drop down combo 
  //for for binding

  //get this machine's host name
  char szHostname[256];
  szHostname[0] = '\0';
  if (gethostname(szHostname, sizeof(szHostname)) != 0)
  {
    TRACE(_T("CPJNSMTPAppConfigurationDlg::AddLocalIpsToBindCombo, Failed in call to gethostname, WSAGetLastError returns %d\n"), WSAGetLastError());
    return FALSE;
  }

  //get host information from the host name
  ATL::CSocketAddr lookup;
  int nError = lookup.FindAddr(CString(szHostname), 80, AI_PASSIVE, 0, 0, 0);
  if (nError != 0)
  {
    TRACE(_T("CPJNSMTPAppConfigurationDlg::AddLocalIpsToBindCombo, Failed in call to gethostbyname, WSAGetLastError returns %d\n"), WSAGetLastError());
    return FALSE;
  }

  ADDRINFOT* const pAddresses = lookup.GetAddrInfoList();
  ASSERT(pAddresses != NULL);


  //Add the addresses to the combo box (Note this app only bother's adding IPv4 addresses)
  m_ctrlIPAddresses.ResetContent();
  m_ctrlIPAddresses.AddString(_T("ANY_IP_ADDRESS"));
  ADDRINFOT* pCurrentAddress = pAddresses;
  while (pCurrentAddress != NULL)
  {
    switch (pCurrentAddress->ai_family)
    {
      case AF_INET:
      {
        CString sIP;
        struct sockaddr_in* sockaddr_ipv4 = reinterpret_cast<sockaddr_in*>(pCurrentAddress->ai_addr);
        InetNtop(AF_INET, &sockaddr_ipv4->sin_addr, sIP.GetBufferSetLength(46), 46);
        sIP.ReleaseBuffer();
        m_ctrlIPAddresses.AddString(sIP);
        break;
      }
      default:
      {
        break;
      }
    }

    //Prepare for the next loop
    pCurrentAddress = pCurrentAddress->ai_next;
  }
  m_ctrlIPAddresses.SetCurSel(0);

  return TRUE;
}

void CPJNSMTPAppConfigurationDlg::OnSelchangeAuthenticate() 
{
  int nAuthentication = m_ctrlAuthenticate.GetCurSel();
  BOOL bAuthenticated = (nAuthentication != CPJNSMTPConnection::AUTH_NONE);
  m_ctrlUsername.EnableWindow(bAuthenticated);
  m_ctrlPromptUsername.EnableWindow(bAuthenticated);
  m_ctrlPassword.EnableWindow(bAuthenticated);
  m_ctrlPromptPassword.EnableWindow(bAuthenticated);
}

void CPJNSMTPAppConfigurationDlg::DDX_CBData(CDataExchange* pDX, int nIDC, DWORD& dwItemData)
{
  HWND hWndCtrl = pDX->PrepareCtrl(nIDC);
  CComboBox* pCombo = static_cast<CComboBox*>(CWnd::FromHandle(hWndCtrl));
  AFXASSUME(pCombo);
  if (pDX->m_bSaveAndValidate)
  {
    dwItemData = 0L;
    int nCurSel = pCombo->GetCurSel();
    if (nCurSel != CB_ERR)
      dwItemData = static_cast<DWORD>(pCombo->GetItemData(nCurSel));
  }
  else
  {
    int nIndex;
    for (nIndex = pCombo->GetCount()-1; nIndex>=0; nIndex--)
    {
      DWORD dwData = static_cast<DWORD>(pCombo->GetItemData(nIndex));
      if (dwData == dwItemData)
      {
        pCombo->SetCurSel(nIndex);
        break;
      }
    }
    if (nIndex < 0) // item wasn't found
      pCombo->SetWindowText(_T("???"));
  }
}

void CPJNSMTPAppConfigurationDlg::OnConnectionType() 
{
  CDataExchange DX(this, TRUE);
  DDX_CBData(&DX, IDC_CONNECTIONTYPE, m_ConnectionType);
  GetDlgItem(IDC_SSL_PROTOCOL)->EnableWindow(m_ConnectionType != CPJNSMTPConnection::PlainText);
}

void CPJNSMTPAppConfigurationDlg::OnDSN() 
{
  CDataExchange DX(this, TRUE);
  DDX_Check(&DX, IDC_DSN, m_bDSN);

  GetDlgItem(IDC_DSN_DELAY)->EnableWindow(m_bDSN);
  GetDlgItem(IDC_DSN_FAILURE)->EnableWindow(m_bDSN);
  GetDlgItem(IDC_DSN_SUCCESS)->EnableWindow(m_bDSN);
  GetDlgItem(IDC_DSN_HEADERS)->EnableWindow(m_bDSN);
}
