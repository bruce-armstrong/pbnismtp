#include "stdafx.h"
#include "PJNSMTPApp.h"
#include "PJNSMTPAppDlg.h"
#include "PJNSMTPAppConfigurationDlg.h"
#include "PJNSMTPApp.h"
#include "SendingDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif //#ifdef _DEBUG

//The class we use to display the progress control dialog
class CPJNSMTPConnectionEx : public CPJNSMTPConnection
{
public:
//Constructors / Destructors
  CPJNSMTPConnectionEx();

//methods
  virtual BOOL OnSendProgress(DWORD dwCurrentBytes, DWORD dwTotalBytes);

//Member variables
  CSendingDlg* m_pProgressDlg;
  BOOL m_bSetRange;
};

//The structure we pass to the worker thread
class CThreadSendData
{ 
public:
  CThreadSendData() : m_pConnection(NULL),
                      m_pProgressDlg(NULL),
                      m_pRecipients(NULL),
                      m_pbyMessage(NULL),
                      m_dwMessageSize(0),
                      m_pFrom(NULL),
                      m_bSentOK(FALSE),
                      m_dwMainThreadID(0),
                      m_ExceptionHr(S_OK)
  {
  }

  CPJNSMTPConnectionEx*           m_pConnection;
  CSendingDlg*                    m_pProgressDlg;
  CPJNSMTPMessage::CAddressArray* m_pRecipients;
  CString                         m_sFilename;
  const BYTE*                     m_pbyMessage;
  DWORD                           m_dwMessageSize;
  CPJNSMTPAddress*                m_pFrom;
  BOOL                            m_bSentOK;
  DWORD                           m_dwMainThreadID;
  CString                         m_sExceptionMessage;
  CString                         m_sLastResponse;
  HRESULT                         m_ExceptionHr;
};


CPJNSMTPConnectionEx::CPJNSMTPConnectionEx() : m_pProgressDlg(NULL),
                                               m_bSetRange(FALSE)
{
}

BOOL CPJNSMTPConnectionEx::OnSendProgress(DWORD dwCurrentBytes, DWORD dwTotalBytes)
{
  if (m_pProgressDlg != NULL)
  {
    //Set the range
    if (!m_bSetRange)
    {
      m_bSetRange = TRUE;
      m_pProgressDlg->m_ctrlProgress.SetRange(0, 100);
    }
    
    //Set the percentage done
    m_pProgressDlg->m_ctrlProgress.SetPos(dwCurrentBytes * 100 / dwTotalBytes);
  
    //Should we continue
    BOOL bContinue = !m_pProgressDlg->AbortFlag();
    if (!bContinue)
      Disconnect(FALSE);

    return bContinue;
  }
  else
    return TRUE;
}


CPJNSMTPAppDlg::CPJNSMTPAppDlg(CWnd* pParent)	: CDialog(CPJNSMTPAppDlg::IDD, pParent),
                                                m_bDNSLookup(FALSE)
{
  //Load in the values from the registry
  CWinApp* pApp = AfxGetApp();
  CString sSection(_T("Settings"));
  m_nPort = pApp->GetProfileInt(sSection, _T("Port"), 25);
  m_sAddress = pApp->GetProfileString(sSection, _T("Address"));
  m_sHost = pApp->GetProfileString(sSection, _T("Host"));
  m_sName = pApp->GetProfileString(sSection, _T("Name"));
  m_Auth = static_cast<CPJNSMTPConnection::AuthenticationMethod>(pApp->GetProfileInt(sSection, _T("Authenticate"), CPJNSMTPConnection::AUTH_NONE));
  
  //Decrypt the username
  BYTE* pEncryptedUsername = NULL;
  UINT nEncryptedUsernameLength = 0;
  if (pApp->GetProfileBinary(sSection, _T("UsernameE"), &pEncryptedUsername, &nEncryptedUsernameLength))
  {
    DATA_BLOB dbIn;
    dbIn.cbData = nEncryptedUsernameLength;
    dbIn.pbData = pEncryptedUsername;
    DATA_BLOB dbOut;
    if (CryptUnprotectData(&dbIn, NULL, NULL, NULL, NULL, CRYPTPROTECT_UI_FORBIDDEN, &dbOut))
    {
      m_sUsername = reinterpret_cast<LPCWSTR>(dbOut.pbData);
      LocalFree(dbOut.pbData);
    }
  }
  delete [] pEncryptedUsername;
  
  //Decrypt the password
  BYTE* pEncryptedPassword = NULL;
  UINT nEncryptedPasswordLength = 0;
  if (pApp->GetProfileBinary(sSection, _T("PasswordE"), &pEncryptedPassword, &nEncryptedPasswordLength))
  {
    DATA_BLOB dbIn;
    dbIn.cbData = nEncryptedPasswordLength;
    dbIn.pbData = pEncryptedPassword;
    DATA_BLOB dbOut;
    if (CryptUnprotectData(&dbIn, NULL, NULL, NULL, NULL, CRYPTPROTECT_UI_FORBIDDEN, &dbOut))
    {
      m_sPassword = reinterpret_cast<LPCWSTR>(dbOut.pbData);
      LocalFree(dbOut.pbData);
    }
  }
  delete [] pEncryptedPassword;
  
  m_bAutoDial = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("Autodial"), FALSE));
  m_sBoundIP = pApp->GetProfileString(sSection, _T("BoundIP"));
  m_sEncodingFriendly = pApp->GetProfileString(sSection, _T("EncodingFriendly"), _T("Western European (ISO)"));
  m_sEncodingCharset = pApp->GetProfileString(sSection, _T("EncodingCharset"), _T("iso-8859-1"));
  m_bMime = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("Mime"), FALSE));
  m_bHTML = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("HTML"), FALSE));
  m_Priority = static_cast<CPJNSMTPMessage::PRIORITY>(pApp->GetProfileInt(sSection, _T("Priority"), CPJNSMTPMessage::NoPriority));
  m_ConnectionType = static_cast<CPJNSMTPConnection::ConnectionType>(pApp->GetProfileInt(sSection, _T("ConnectionType"), CPJNSMTPConnection::PlainText));
  m_SSLProtocol = static_cast<CPJNSMTPConnection::SSLProtocol>(pApp->GetProfileInt(sSection, _T("SSLProtocol"), CPJNSMTPConnection::OSDefault));
  m_bDSN = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("DSN"), FALSE));
  m_bDSNSuccess = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("DSNSuccess"), FALSE));
  m_bDSNFailure = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("DSNFailure"), FALSE));
  m_bDSNDelay = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("DSNDelay"), FALSE));
  m_bDSNHeaders = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("DSNHeaders"), TRUE));
  m_bMDN = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("MDN"), FALSE));
  m_bInitialClientResponse = static_cast<BOOL>(pApp->GetProfileInt(sSection, _T("InitialClientResponse"), FALSE));

#ifdef _DEBUG
  /* Some test code to exercise SMTP body part creation
  CPJNSMTPBodyPart asciiBodyPart;
  asciiBodyPart.SetCharset(_T("us-ascii"));
  asciiBodyPart.SetQuotedPrintable(FALSE);
  asciiBodyPart.SetText(_T("This is the plain text body part"));

  CPJNSMTPBodyPart HTMLBodyPart;
  HTMLBodyPart.SetContentType(_T("text/html"));
  HTMLBodyPart.SetCharset(_T("us-ascii"));
  HTMLBodyPart.SetQuotedPrintable(FALSE);
  HTMLBodyPart.SetRawBody("<html><head></head><body>This is the <b>HTML</b> body part</body></html>");

  CPJNSMTPBodyPart alternativeBodyPart;
  alternativeBodyPart.SetContentType(_T("multipart/alternative"));
  alternativeBodyPart.AddChildBodyPart(asciiBodyPart);
  alternativeBodyPart.AddChildBodyPart(HTMLBodyPart);

  CPJNSMTPBodyPart attachmentBodyPart;
  attachmentBodyPart.SetContentType(_T("text/plain"));
  attachmentBodyPart.SetAttachment(TRUE);
  attachmentBodyPart.SetBase64(FALSE);
  attachmentBodyPart.SetQuotedPrintable(FALSE);
  attachmentBodyPart.SetTitle(_T("bla.txt"));
  LPCSTR pszAttachment = "This text should appear in the attachment";
  attachmentBodyPart.SetRawBody(pszAttachment);

  CPJNSMTPMessage message;
  message.SetMime(TRUE);
  message.m_RootPart.AddChildBodyPart(alternativeBodyPart);
  message.m_RootPart.AddChildBodyPart(attachmentBodyPart);
  message.m_From = CPJNSMTPAddress(_T("somebody@somedomain.com"));
  message.m_sSubject = _T("This is the subject of the email");
  message.SaveToDisk(_T("d:\\temp\\bla.eml"));
  */
#endif //#ifdef _DEBUG

  m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CPJNSMTPAppDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);

  DDX_Control(pDX, IDC_BODY, m_wndBody);
  DDX_Text(pDX, IDC_BCC, m_sBCC);
  DDX_Text(pDX, IDC_BODY, m_sBody);
  DDX_Text(pDX, IDC_CC, m_sCC);
  DDX_Text(pDX, IDC_FILE, m_sFile);
  DDX_Text(pDX, IDC_SUBJECT, m_sSubject);
  DDX_Text(pDX, IDC_TO, m_sTo);
  DDX_Check(pDX, IDC_DNS_LOOKUP, m_bDNSLookup);

  if (pDX->m_bSaveAndValidate)
  {
    if (m_sTo.IsEmpty())
    {
      AfxMessageBox(_T("Please specify a valid To address"), MB_OK | MB_ICONEXCLAMATION);
      pDX->PrepareEditCtrl(IDC_TO);
      pDX->Fail();
    }
  }
}

BEGIN_MESSAGE_MAP(CPJNSMTPAppDlg, CDialog)
  ON_WM_PAINT()
  ON_WM_QUERYDRAGICON()
  ON_BN_CLICKED(IDC_CONFIGURATION, &CPJNSMTPAppDlg::OnConfiguration)
  ON_BN_CLICKED(IDC_SEND, &CPJNSMTPAppDlg::OnSend)
  ON_BN_CLICKED(IDC_BROWSE_FILE, &CPJNSMTPAppDlg::OnBrowseFile)
  ON_WM_DESTROY()
  ON_BN_CLICKED(IDC_SEND_FROM_DISK, &CPJNSMTPAppDlg::OnSendFromDisk)
  ON_BN_CLICKED(IDC_SEND_FROM_DISK_WITH_UI, &CPJNSMTPAppDlg::OnSendFromDiskWithUI)
  ON_BN_CLICKED(IDC_SEND_FROM_MEMORY, &CPJNSMTPAppDlg::OnSendFromMemory)
  ON_BN_CLICKED(IDC_SEND_FROM_MEMORY_WITH_UI, &CPJNSMTPAppDlg::OnSendFromMemoryWithUI)
  ON_BN_CLICKED(IDC_SAVE_TO_DISK, &CPJNSMTPAppDlg::OnSaveToDisk)
END_MESSAGE_MAP()

BOOL CPJNSMTPAppDlg::OnInitDialog()
{
  //Let the parent class do its thing
  CDialog::OnInitDialog();

  //Set big icon
  SetIcon(m_hIcon, TRUE);	
  
  //allow as much text as possible in the body edit control
  m_wndBody.SetLimitText(0); 

  return TRUE;
}

void CPJNSMTPAppDlg::OnPaint() 
{
  if (IsIconic())
  {
    CPaintDC dc(this); //device context for painting

    SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

    //Center icon in client rectangle
    int cxIcon = GetSystemMetrics(SM_CXICON);
    int cyIcon = GetSystemMetrics(SM_CYICON);
    CRect rect;
    GetClientRect(&rect);
    int x = (rect.Width() - cxIcon + 1) / 2;
    int y = (rect.Height() - cyIcon + 1) / 2;

    //Draw the icon
    dc.DrawIcon(x, y, m_hIcon);
  }
  else
    CDialog::OnPaint();
}

HCURSOR CPJNSMTPAppDlg::OnQueryDragIcon()
{
  return static_cast<HCURSOR>(m_hIcon);
}

void CPJNSMTPAppDlg::OnConfiguration() 
{
  CPJNSMTPAppConfigurationDlg dlg;
  dlg.m_sAddress               = m_sAddress;
  dlg.m_sHost                  = m_sHost;
  dlg.m_sName                  = m_sName;
  dlg.m_nPort                  = m_nPort;
  dlg.m_Auth                   = m_Auth;
  dlg.m_sUsername              = m_sUsername;
  dlg.m_sPassword              = m_sPassword;
  dlg.m_bAutoDial              = m_bAutoDial;
  dlg.m_sBoundIP               = m_sBoundIP;
  dlg.m_sEncodingCharset       = m_sEncodingCharset;
  dlg.m_sEncodingFriendly      = m_sEncodingFriendly;
  dlg.m_bMime                  = m_bMime;
  dlg.m_bHTML                  = m_bHTML;
  dlg.m_Priority               = m_Priority;
  dlg.m_ConnectionType         = m_ConnectionType;
  dlg.m_SSLProtocol            = m_SSLProtocol;
  dlg.m_bDSN                   = m_bDSN;
  dlg.m_bDSNSuccess            = m_bDSNSuccess;
  dlg.m_bDSNFailure            = m_bDSNFailure;
  dlg.m_bDSNDelay              = m_bDSNDelay;
  dlg.m_bDSNHeaders            = m_bDSNHeaders;
  dlg.m_bMDN                   = m_bMDN;
  dlg.m_bInitialClientResponse = m_bInitialClientResponse;
  if (dlg.DoModal() == IDOK)
  {
    m_sAddress               = dlg.m_sAddress;
    m_sHost                  = dlg.m_sHost;
    m_sName                  = dlg.m_sName;
    m_nPort                  = dlg.m_nPort;
    m_Auth                   = static_cast<CPJNSMTPConnection::AuthenticationMethod>(dlg.m_Auth);
    m_sUsername              = dlg.m_sUsername;
    m_sPassword              = dlg.m_sPassword;
    m_bAutoDial              = dlg.m_bAutoDial;
    m_sBoundIP               = dlg.m_sBoundIP;
    m_sEncodingCharset       = dlg.m_sEncodingCharset;
    m_sEncodingFriendly      = dlg.m_sEncodingFriendly;
    m_bMime                  = dlg.m_bMime;
    m_bHTML                  = dlg.m_bHTML;
    m_Priority               = static_cast<CPJNSMTPMessage::PRIORITY>(dlg.m_Priority);
    m_ConnectionType         = static_cast<CPJNSMTPConnection::ConnectionType>(dlg.m_ConnectionType);
    m_SSLProtocol            = static_cast<CPJNSMTPConnection::SSLProtocol>(dlg.m_SSLProtocol);
    m_bDSN                   = dlg.m_bDSN;
    m_bDSNSuccess            = dlg.m_bDSNSuccess;
    m_bDSNFailure            = dlg.m_bDSNFailure;
    m_bDSNDelay              = dlg.m_bDSNDelay;
    m_bDSNHeaders            = dlg.m_bDSNHeaders;
    m_bMDN                   = dlg.m_bMDN;
    m_bInitialClientResponse = dlg.m_bInitialClientResponse;
  }
}

CPJNSMTPMessage* CPJNSMTPAppDlg::CreateMessage()
{
  //Create the message
  CPJNSMTPMessage* pMessage = new CPJNSMTPMessage;
  CPJNSMTPBodyPart attachment;

  //Set the mime flag
  pMessage->SetMime(m_bMime);

  //Set the charset of the message and all attachments
  pMessage->SetCharset(m_sEncodingCharset);
  attachment.SetCharset(m_sEncodingCharset);

  //Set the message priority
  pMessage->m_Priority = m_Priority;

  //Setup the all the recipient types for this message
  pMessage->ParseMultipleRecipients(m_sTo, pMessage->m_To);
  if (!m_sCC.IsEmpty()) 
    pMessage->ParseMultipleRecipients(m_sCC, pMessage->m_CC);
  if (!m_sBCC.IsEmpty()) 
    pMessage->ParseMultipleRecipients(m_sBCC, pMessage->m_BCC);
  if (!m_sSubject.IsEmpty()) 
    pMessage->m_sSubject = m_sSubject;
  if (!m_sBody.IsEmpty())
  {
    if (m_bHTML)
      pMessage->AddHTMLBody(m_sBody);
    else
      pMessage->AddTextBody(m_sBody);
  }

  //Add the attachment(s) if necessary
  if (!m_sFile.IsEmpty()) 
    pMessage->AddMultipleAttachments(m_sFile);		

  //Setup the from address
  if (m_sName.IsEmpty()) 
    pMessage->m_From = CPJNSMTPAddress(m_sAddress);
  else 
    pMessage->m_From = CPJNSMTPAddress(m_sName, m_sAddress);

#ifndef _DEBUG
#ifdef CPJNSMTP_MFC_EXTENSIONS
  pMessage->m_sXMailer.Empty(); //Release builds do not include a XMailer header
#else
  pMessage->m_sXMailer.clear(); //Release builds do not include a XMailer header
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
#endif //#ifndef _DEBUG

  if (m_bDSN)
  {
    pMessage->m_DSNReturnType = m_bDSNHeaders ? CPJNSMTPMessage::HeadersOnly : CPJNSMTPMessage::FullEmail;
    pMessage->m_dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED;
    if (m_bDSNSuccess)
      pMessage->m_dwDSN |= CPJNSMTPMessage::DSN_SUCCESS;
    if (m_bDSNFailure)
      pMessage->m_dwDSN |= CPJNSMTPMessage::DSN_FAILURE;
    if (m_bDSNDelay)
      pMessage->m_dwDSN |= CPJNSMTPMessage::DSN_DELAY;
  }
  else
    pMessage->m_dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED;
    
  if (m_bMDN)
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    pMessage->m_MessageDispositionEmailAddresses.Add(m_sAddress);
  #else
    pMessage->m_MessageDispositionEmailAddresses.push_back(CPJNSMTPString(m_sAddress));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }

#ifdef _DEBUG
  //Add one custom header (for test purpose)
#ifdef CPJNSMTP_MFC_EXTENSIONS
  pMessage->m_CustomHeaders.Add(_T("X-Program: CSTMPMessageTester"));
#else
  pMessage->m_CustomHeaders.push_back(_T("X-Program: CSTMPMessageTester"));
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //Try out the copy constructor and operator= methods
  CPJNSMTPMessage copyOfMessage(*pMessage);
#endif //#ifdef _DEBUG

  return pMessage;
}
  

void CPJNSMTPAppDlg::OnSend() 
{
  if (UpdateData(TRUE))
  {
    if ((m_sHost.IsEmpty() && !m_bDNSLookup) || m_sAddress.IsEmpty()) 
    {
      AfxMessageBox(_T("You need to configure your settings first"));
      OnConfiguration();
    }
    else 
    {
      //Display a wait cursor
      CWaitCursor wc;

      CPJNSMTPMessage* pMessage = NULL;
      BOOL bCloseGracefully = TRUE;
      try
      {
        //Create the message
        pMessage = CreateMessage();

        CPJNSMTPString sHost;
        BOOL bSend = TRUE;
        if (m_bDNSLookup)
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          if (pMessage->m_To.GetSize() == 0)
        #else
          if (pMessage->m_To.size() == 0)
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          {
            AfxMessageBox(_T("At least one receipient must be specified to use the DNS lookup option"));
            bSend = FALSE;
          }
          else
          {
            CPJNSMTPString sAddress(pMessage->m_To[0].m_sEmailAddress);
          #ifdef CPJNSMTP_MFC_EXTENSIONS
            int nAmpersand = sAddress.Find(_T("@"));
            if (nAmpersand == -1)
          #else
            CPJNSMTPString::size_type nAmpersand = sAddress.find(_T("@"));
            if (nAmpersand == CPJNSMTPString::npos)
          #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
            {
              CString sMsg;
            #ifdef CPJNSMTP_MFC_EXTENSIONS
              sMsg.Format(_T("Unable to determine the domain for the email address \"%s\""), sAddress.operator LPCTSTR());
            #else
              sMsg.Format(_T("Unable to determine the domain for the email address \"%s\""), sAddress.c_str());
            #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
              AfxMessageBox(sMsg);
              bSend = FALSE;
            }
            else
            {
              //We just pick the first MX record found, other implementations could ask the user
              //or automatically pick the lowest priority record
            #ifdef CPJNSMTP_MFC_EXTENSIONS
              CPJNSMTPString sDomain(sAddress.Mid(nAmpersand + 1));
            #else
              CPJNSMTPString sDomain(sAddress.substr(nAmpersand + 1));
            #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
              CPJNSMTPStringArray servers;
              CPJNSMTPWordArray priorities;
            #ifdef CPJNSMTP_MFC_EXTENSIONS
              if (!m_Connection.MXLookup(sDomain, servers, priorities))
            #else
              if (!m_Connection.MXLookup(sDomain.c_str(), servers, priorities))
            #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
              {
                CString sMsg;
              #ifdef CPJNSMTP_MFC_EXTENSIONS
                sMsg.Format(_T("Unable to perform a DNS MX lookup for the domain \"%s\", Error Code:%u"), sDomain.operator LPCTSTR(), GetLastError());
              #else
                sMsg.Format(_T("Unable to perform a DNS MX lookup for the domain \"%s\", Error Code:%u"), sDomain.c_str(), GetLastError());
              #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
                AfxMessageBox(sMsg);
                bSend = FALSE;
              }
              else
              {
                sHost = servers[0];
              }
            }
          }
        }
        else
          sHost = m_sHost;

        //Instantiate the class which does the work for us
        if (bSend)
        {
          //Disconnect if already connected
          if (m_Connection.IsConnected())
            m_Connection.Disconnect();

          //Auto connect to the internet?
          if (m_bAutoDial)
            m_Connection.ConnectToInternet();

          //Connect and send the message
          m_Connection.SetSSLProtocol(m_SSLProtocol);
          m_Connection.SetBindAddress(m_sBoundIP);
          m_Connection.SetInitialClientResponse(m_bInitialClientResponse);
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          m_Connection.Connect(sHost, m_Auth, m_sUsername, m_sPassword, m_nPort, m_ConnectionType, pMessage->m_dwDSN);
        #else
          m_Connection.Connect(sHost.c_str(), m_Auth, m_sUsername, m_sPassword, m_nPort, m_ConnectionType, pMessage->m_dwDSN);
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          m_Connection.SendMessage(*pMessage);

          //Auto disconnect from the internet
          if (m_bAutoDial)
            m_Connection.CloseInternetConnection();
        }
      }
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      catch(CPJNSMTPException* pEx)
      {
        //Display the error
        CString sMsg;
        sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR(), pEx->m_sLastResponse.operator LPCTSTR());
        AfxMessageBox(sMsg, MB_ICONSTOP);
        pEx->Delete();
        bCloseGracefully = FALSE;
      }
    #else
      catch(CPJNSMTPException& e)
      {
        //Display the error
        CString sMsg;
        TCHAR szError[4096];
        szError[0] = _T('\0');
        e.GetErrorMessage(szError, _countof(szError));
        sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), e.m_hr, szError, e.m_sLastResponse.c_str());
        AfxMessageBox(sMsg, MB_ICONSTOP);
        bCloseGracefully = FALSE;
      }
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

      //Tidy up the heap memory
      if (pMessage)
        delete pMessage;

      //Disconnect from the server
      try
      {
        //Disconnect if already connected
        if (m_Connection.IsConnected())
          m_Connection.Disconnect(bCloseGracefully);
      }
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      catch(CPJNSMTPException* pEx)
      {
        pEx->Delete();
      }
    #else
      catch(CPJNSMTPException& /*e*/)
      {
      }
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
  }
}

void CPJNSMTPAppDlg::OnBrowseFile() 
{
  CDataExchange DX(this, TRUE);
  DDX_Text(&DX, IDC_FILE, m_sFile);
  CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("All Files (*.*)|*.*||"));
  if (dlg.DoModal() == IDOK) 
  {
    CString sNewFile(dlg.GetPathName());
    if (m_sFile.GetLength())
    {
      m_sFile += _T(", ");
      m_sFile += sNewFile;
    }
    else
      m_sFile = sNewFile;

    //Update the UI  
    CDataExchange DX2(this, FALSE);
    DDX_Text(&DX2, IDC_FILE, m_sFile);
  }
}

void CPJNSMTPAppDlg::OnDestroy() 
{
  //Let the parent class do its thing
  CDialog::OnDestroy();
  
  //Save in the values from the registry
  CWinApp* pApp = AfxGetApp();
  CString sSection(_T("Settings"));
  pApp->WriteProfileInt(sSection, _T("Port"), m_nPort);
  pApp->WriteProfileString(sSection, _T("Address"), m_sAddress);
  pApp->WriteProfileString(sSection, _T("Host"), m_sHost);
  pApp->WriteProfileString(sSection, _T("Name"), m_sName);
  pApp->WriteProfileInt(sSection, _T("Authenticate"), m_Auth);
  
  //Encrypt the username
  CStringW sTempUsername(m_sUsername);
  DATA_BLOB dbIn;
  int nUsernameLength = sTempUsername.GetLength();
  dbIn.cbData = (nUsernameLength + 1) * sizeof(wchar_t);
  dbIn.pbData = reinterpret_cast<BYTE*>(sTempUsername.GetBuffer(nUsernameLength));
  DATA_BLOB dbOut;
  if (CryptProtectData(&dbIn, NULL, NULL, NULL, NULL, CRYPTPROTECT_UI_FORBIDDEN, &dbOut))
  {
    pApp->WriteProfileBinary(sSection, _T("UsernameE"), dbOut.pbData, dbOut.cbData);
    LocalFree(dbOut.pbData);
  }
  sTempUsername.ReleaseBuffer();
  
  //Encrypt the password
  CStringW sTempPassword(m_sPassword);
  int nPasswordLength = sTempPassword.GetLength();
  dbIn.cbData = (nPasswordLength + 1) * sizeof(wchar_t);
  dbIn.pbData = reinterpret_cast<BYTE*>(sTempPassword.GetBuffer(nPasswordLength));
  if (CryptProtectData(&dbIn, NULL, NULL, NULL, NULL, CRYPTPROTECT_UI_FORBIDDEN, &dbOut))
  {
    pApp->WriteProfileBinary(sSection, _T("PasswordE"), dbOut.pbData, dbOut.cbData);
    LocalFree(dbOut.pbData);
  }
  sTempPassword.ReleaseBuffer();
  
  pApp->WriteProfileInt(sSection, _T("Autodial"), m_bAutoDial);
  pApp->WriteProfileString(sSection, _T("BoundIP"), m_sBoundIP);
  pApp->WriteProfileString(sSection, _T("EncodingFriendly"), m_sEncodingFriendly);
  pApp->WriteProfileString(sSection, _T("EncodingCharset"), m_sEncodingCharset);
  pApp->WriteProfileInt(sSection, _T("Mime"), m_bMime);
  pApp->WriteProfileInt(sSection, _T("HTML"), m_bHTML);
  pApp->WriteProfileInt(sSection, _T("Priority"), m_Priority);
  pApp->WriteProfileInt(sSection, _T("ConnectionType"), m_ConnectionType);
  pApp->WriteProfileInt(sSection, _T("SSLProtocol"), m_SSLProtocol);
  pApp->WriteProfileInt(sSection, _T("DSN"), m_bDSN);
  pApp->WriteProfileInt(sSection, _T("DSNSuccess"), m_bDSNSuccess);
  pApp->WriteProfileInt(sSection, _T("DSNFailure"), m_bDSNFailure);
  pApp->WriteProfileInt(sSection, _T("DSNDelay"), m_bDSNDelay);
  pApp->WriteProfileInt(sSection, _T("DSNHeaders"), m_bDSNHeaders);
  pApp->WriteProfileInt(sSection, _T("MDN"), m_bMDN);
  pApp->WriteProfileInt(sSection, _T("InitialClientResponse"), m_bInitialClientResponse);
}

void CPJNSMTPAppDlg::OnSendFromDiskWithUI()
{
  SendFromDisk(TRUE);
}

void CPJNSMTPAppDlg::OnSendFromDisk()
{
  SendFromDisk(FALSE);
}

void CPJNSMTPAppDlg::OnSendFromMemoryWithUI()
{
  SendFromMemory(TRUE);
}

void CPJNSMTPAppDlg::OnSendFromMemory()
{
  SendFromMemory(FALSE);
}

void CPJNSMTPAppDlg::SendFromDisk(BOOL bUI)
{
  if (UpdateData(TRUE))
  {
    if (m_sHost.IsEmpty() || m_sAddress.IsEmpty()) 
    {
      AfxMessageBox(_T("You need to configure your settings first"));
      OnConfiguration();
    }
    else
    {
      //Bring up a standard File Open dialog 
      CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("Email Files (*.eml)|*.eml|All Files (*.*)|*.*||"), NULL);
      if (dlg.DoModal() == IDOK)
      {
        //Display a wait cursor
        CWaitCursor wait;

        //Instantiate the class which does the work for us
        CPJNSMTPConnectionEx connection;

        //Auto connect to the internet?
        if (m_bAutoDial)
          connection.ConnectToInternet();

        //Connect and send the message  
        try
        {
          connection.SetBindAddress(m_sBoundIP);
          connection.Connect(m_sHost, m_Auth, m_sUsername, m_sPassword, m_nPort, m_ConnectionType);

          //Get the from field
          CPJNSMTPAddress from(m_sName, m_sAddress);

          //Get the recipients (for testing purposes we only bother looking at m_sTo)
          CPJNSMTPMessage::CAddressArray Recipients;
          CPJNSMTPMessage::ParseMultipleRecipients(m_sTo, Recipients);

          //Send the message
          if (!bUI)
          {
            CPJNSMTPAsciiString sENVID;
            connection.SendMessage(dlg.GetPathName(), Recipients, from, sENVID);
          }
          else
          {
            //The progress dialog instance
            CSendingDlg dlg2;

            //Let the connection class know about the progress dialog
            connection.m_pProgressDlg = &dlg2;

            //Fill in the structure we need to pass to the worker function
            CThreadSendData tsd;
            tsd.m_pConnection = &connection;
            tsd.m_pProgressDlg = &dlg2;
            tsd.m_pRecipients = &Recipients;
            tsd.m_pFrom = &from;
            tsd.m_dwMainThreadID = GetCurrentThreadId();
            tsd.m_sFilename = dlg.GetPathName();

            //Create the worker thread which will do the actual sending
            CWinThread* pWorkerThread = AfxBeginThread(_SendingWorkerThread, &tsd, THREAD_PRIORITY_NORMAL, CREATE_SUSPENDED);
            pWorkerThread->m_bAutoDelete = FALSE;
            pWorkerThread->ResumeThread();

            //Bring up the progress dialog
            dlg2.DoModal();

            //Wait for the thread to close down
            WaitForSingleObject(pWorkerThread->m_hThread, INFINITE);
            delete pWorkerThread;

            if (!tsd.m_bSentOK)
            {
              CString sMsg;
              sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), tsd.m_ExceptionHr, tsd.m_sExceptionMessage.operator LPCTSTR(), tsd.m_sLastResponse.operator LPCTSTR());
              AfxMessageBox(sMsg, MB_ICONSTOP);
            }
          }
        }
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        catch(CPJNSMTPException* pEx)
        {
          //Display the error
          CString sMsg;
          sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR(), pEx->m_sLastResponse.operator LPCTSTR());
          AfxMessageBox(sMsg, MB_ICONSTOP);
  
          pEx->Delete();
        }
      #else
        catch(CPJNSMTPException& e)
        {
          //Display the error
          TCHAR szError[4096];
          szError[0] = _T('\0');
          e.GetErrorMessage(szError, _countof(szError));
          CString sMsg;
          sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), e.m_hr, szError, e.m_sLastResponse.c_str());
          AfxMessageBox(sMsg, MB_ICONSTOP);
        }
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

        //Auto disconnect from the internet
        if (m_bAutoDial)
          connection.CloseInternetConnection();
      }	
    }
  }
}

void CPJNSMTPAppDlg::OnSaveToDisk() 
{
  if (UpdateData(TRUE))
  {
    //Bring up a standard File Save dialog 
    CFileDialog dlg(FALSE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("Email Files (*.eml)|*.eml|All Files (*.*)|*.*||"), NULL);
    if (dlg.DoModal() == IDOK)
    {
      //And save it to disk
      CPJNSMTPMessage* pMessage = CreateMessage();
      try
      {
        pMessage->SaveToDisk(dlg.GetPathName());
      }
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      catch(CPJNSMTPException* pEx)
      {
        CString sMsg;
        sMsg.Format(_T("Failed to save message to disk, Error:%08X, %s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR());
        AfxMessageBox(sMsg, MB_ICONSTOP);

        pEx->Delete();
      }
    #else
      catch(CPJNSMTPException& e)
      {
        TCHAR szError[4096];
        szError[0] = _T('\0');
        e.GetErrorMessage(szError, _countof(szError));
        CString sMsg;
        sMsg.Format(_T("Failed to save message to disk, Error:%08X, %s"), e.m_hr, szError);
        AfxMessageBox(sMsg, MB_ICONSTOP);
      }
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

      //Tidy up the heap memory
      delete pMessage;
    }	
  }
}

UINT CPJNSMTPAppDlg::_SendingWorkerThread(LPVOID lpParam)
{
  //Validate our parameters
  CThreadSendData* pTSD = static_cast<CThreadSendData*>(lpParam);
  AFXASSUME(pTSD);
  AFXASSUME(pTSD->m_pConnection != NULL);
  AFXASSUME(pTSD->m_pRecipients != NULL);
  AFXASSUME(pTSD->m_pFrom != NULL);
  ASSERT(pTSD->m_pProgressDlg != NULL);

  AttachThreadInput(pTSD->m_dwMainThreadID, GetCurrentThreadId(), TRUE);

  //Wait until the dialog has been created before we continue
  WaitForSingleObject(pTSD->m_pProgressDlg->m_evtInitialized, INFINITE);

  //Do the actual sending
  try
  {
    CPJNSMTPAsciiString sENVID;
    if (pTSD->m_sFilename.GetLength())
      pTSD->m_pConnection->SendMessage(pTSD->m_sFilename, *pTSD->m_pRecipients, *pTSD->m_pFrom, sENVID);
    else
      pTSD->m_pConnection->SendMessage(pTSD->m_pbyMessage, pTSD->m_dwMessageSize, *pTSD->m_pRecipients, *pTSD->m_pFrom, sENVID);
    pTSD->m_bSentOK = !pTSD->m_pProgressDlg->AbortFlag();
  }
#ifdef CPJNSMTP_MFC_EXTENSIONS
  catch(CPJNSMTPException* pEx)
  {
    pTSD->m_ExceptionHr = pEx->m_hr;
    pTSD->m_sExceptionMessage = pEx->GetErrorMessage();
    pTSD->m_sLastResponse = pEx->m_sLastResponse;
    pEx->Delete();
    pTSD->m_bSentOK = FALSE;
  }
#else
  catch(CPJNSMTPException& e)
  {
    pTSD->m_ExceptionHr = e.m_hr;
    TCHAR szError[4096];
    szError[0] = _T('\0');
    e.GetErrorMessage(szError, _countof(szError));
    pTSD->m_sExceptionMessage = szError;
    pTSD->m_sLastResponse = e.m_sLastResponse.c_str();
    pTSD->m_bSentOK = FALSE;
  }
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  pTSD->m_pProgressDlg->SetSafeToClose();
  pTSD->m_pProgressDlg->PostMessage(WM_CLOSE);

  return 0L;
}

void CPJNSMTPAppDlg::SendFromMemory(BOOL bUI)
{
  if (UpdateData(TRUE))
  {
    if (m_sHost.IsEmpty() || m_sAddress.IsEmpty())
    {
      AfxMessageBox(_T("You need to configure your settings first"));
      OnConfiguration();
    }
    else
    {
      //Bring up a standard File Open dialog 
      CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("Email Files (*.eml)|*.eml|All Files (*.*)|*.*||"), NULL);
      if (dlg.DoModal() == IDOK)
      {
        //Display a wait cursor
        CWaitCursor wait;

        //Open up the file we want to send
        ATL::CAtlFile file;
        CString sFile(dlg.GetPathName());
        HRESULT hr = file.Create(sFile, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
        if (FAILED(hr))
        {
          CString sMsg;
          sMsg.Format(_T("Failed to open the file:\'%s\', Error:%08X"), sFile.operator LPCTSTR(), hr);
          AfxMessageBox(sMsg);
          return;
        }

        //Get the length of the file (we only support sending files less than 4GB!)
        ULONGLONG nFileSize = 0;
        hr = file.GetSize(nFileSize);
        if (FAILED(hr))
        {
          CString sMsg;
          sMsg.Format(_T("Failed to obtain the file length for the file:\'%s\', Error:%08X"), sFile.operator LPCTSTR(), hr);
          AfxMessageBox(sMsg);
          return;
        }
        if (nFileSize >= ULONG_MAX)
        {
          CString sMsg;
          sMsg.Format(_T("The file \'%s\' is too large to send"), sFile.operator LPCTSTR());
          AfxMessageBox(sMsg);
          return;
        }
        if (nFileSize == 0)
        {
          CString sMsg;
          sMsg.Format(_T("The file \'%s\' does not contain any data to send"), sFile.operator LPCTSTR());
          AfxMessageBox(sMsg);
          return;
        }

        //Allocate some memory to read in the file
        ATL::CHeapPtr<BYTE> byMessage;
        if (!byMessage.Allocate(static_cast<size_t>(nFileSize)))
        {
          AfxMessageBox(_T("Unable to allocate memory to read from the specified file"));
          return;
        }

        //Read in the file contents
        hr = file.Read(byMessage.m_pData, static_cast<DWORD>(nFileSize));
        if (FAILED(hr))
        {
          CString sMsg;
          sMsg.Format(_T("Failed to read from the file:%s, Error:%08X"), sFile.operator LPCTSTR(), hr);
          AfxMessageBox(sMsg);
          return;
        }

        //Instantiate the class which does the work for us
        CPJNSMTPConnectionEx connection;

        //Auto connect to the internet?
        if (m_bAutoDial)
          connection.ConnectToInternet();

        //Connect and send the message  
        try
        {
          connection.SetBindAddress(m_sBoundIP);
          connection.Connect(m_sHost, m_Auth, m_sUsername, m_sPassword, m_nPort, m_ConnectionType);

          //Get the from field
          CPJNSMTPAddress from(m_sName, m_sAddress);

          //Get the recipients (for testing purposes we only bother looking at m_sTo)
          CPJNSMTPMessage::CAddressArray Recipients;
          CPJNSMTPMessage::ParseMultipleRecipients(m_sTo, Recipients);

          //Send the message
          if (!bUI)
          {
            CPJNSMTPAsciiString sENVID;
            connection.SendMessage(byMessage.m_pData, static_cast<DWORD>(nFileSize), Recipients, from, sENVID);
          }
          else
          {
            //The progress dialog instance
            CSendingDlg dlg2;

            //Let the connection class know about the progress dialog
            connection.m_pProgressDlg = &dlg2;

            //Fill in the structure we need to pass to the worker function
            CThreadSendData tsd;
            tsd.m_pConnection = &connection;
            tsd.m_pProgressDlg = &dlg2;
            tsd.m_pRecipients = &Recipients;
            tsd.m_pFrom = &from;
            tsd.m_dwMainThreadID = GetCurrentThreadId();
            tsd.m_pbyMessage = byMessage.m_pData;
            tsd.m_dwMessageSize = static_cast<DWORD>(nFileSize);

            //Create the worker thread which will do the actual sending
            CWinThread* pWorkerThread = AfxBeginThread(_SendingWorkerThread, &tsd, THREAD_PRIORITY_NORMAL, CREATE_SUSPENDED);
            pWorkerThread->m_bAutoDelete = FALSE;
            pWorkerThread->ResumeThread();

            //Bring up the progress dialog
            dlg2.DoModal();

            //Wait for the thread to close down
            WaitForSingleObject(pWorkerThread->m_hThread, INFINITE);
            delete pWorkerThread;

            if (!tsd.m_bSentOK)
            {
              CString sMsg;
              sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), tsd.m_ExceptionHr, tsd.m_sExceptionMessage.operator LPCTSTR(), tsd.m_sLastResponse.operator LPCTSTR());
              AfxMessageBox(sMsg, MB_ICONSTOP);
            }
          }
        }
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        catch(CPJNSMTPException* pEx)
        {
          //Display the error
          CString sMsg;
          sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR(), pEx->m_sLastResponse.operator LPCTSTR());
          AfxMessageBox(sMsg, MB_ICONSTOP);

          pEx->Delete();
        }
      #else
        catch (CPJNSMTPException& e)
        {
          //Display the error
          TCHAR szError[4096];
          szError[0] = _T('\0');
          e.GetErrorMessage(szError, _countof(szError));
          CString sMsg;
          sMsg.Format(_T("An error occurred sending the message, Error:%08X, %s, Last Response:%s"), e.m_hr, szError, e.m_sLastResponse.c_str());
          AfxMessageBox(sMsg, MB_ICONSTOP);
        }
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

        //Auto disconnect from the internet
        if (m_bAutoDial)
          connection.CloseInternetConnection();
      }
    }
  }
}
