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

//The structure we pass wo the worker thread
struct ThreadSendData
{ 
  CPJNSMTPConnectionEx*           pConnection;
  CSendingDlg*                    pProgressDlg;
  CPJNSMTPMessage::CAddressArray* pRecipients;
  CString                         sFilename;
  CPJNSMTPAddress*                pFrom;
  BOOL                            bSentOK;
  DWORD                           dwMainThreadID;
  CString                         sExceptionMessage;
  HRESULT                         ExceptionHr;
};


CPJNSMTPConnectionEx::CPJNSMTPConnectionEx() : m_pProgressDlg(NULL),
                                               m_bSetRange(FALSE)
{
}

BOOL CPJNSMTPConnectionEx::OnSendProgress(DWORD dwCurrentBytes, DWORD dwTotalBytes)
{
  if (m_pProgressDlg)
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
                                                m_bDirectly(FALSE),
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
  DDX_Check(pDX, IDC_DIRECTLY, m_bDirectly);
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
  ON_BN_CLICKED(IDC_SAVE, &CPJNSMTPAppDlg::OnSave)
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
  dlg.m_sAddress          = m_sAddress;
  dlg.m_sHost             = m_sHost;
  dlg.m_sName             = m_sName;
  dlg.m_nPort             = m_nPort;
  dlg.m_Auth              = m_Auth;
  dlg.m_sUsername         = m_sUsername;
  dlg.m_sPassword         = m_sPassword;
  dlg.m_bAutoDial         = m_bAutoDial;
  dlg.m_sBoundIP          = m_sBoundIP;
  dlg.m_sEncodingCharset  = m_sEncodingCharset;
  dlg.m_sEncodingFriendly = m_sEncodingFriendly;
  dlg.m_bMime             = m_bMime;
  dlg.m_bHTML             = m_bHTML;
  dlg.m_Priority          = m_Priority;
  dlg.m_ConnectionType    = m_ConnectionType;
  dlg.m_SSLProtocol       = m_SSLProtocol;
  dlg.m_bDSN              = m_bDSN;
  dlg.m_bDSNSuccess       = m_bDSNSuccess;
  dlg.m_bDSNFailure       = m_bDSNFailure;
  dlg.m_bDSNDelay         = m_bDSNDelay;
  dlg.m_bDSNHeaders       = m_bDSNHeaders;
  dlg.m_bMDN              = m_bMDN;
  if (dlg.DoModal() == IDOK)
  {
    m_sAddress          = dlg.m_sAddress;
    m_sHost             = dlg.m_sHost;
    m_sName             = dlg.m_sName;
    m_nPort             = dlg.m_nPort;
    m_Auth              = static_cast<CPJNSMTPConnection::AuthenticationMethod>(dlg.m_Auth);
    m_sUsername         = dlg.m_sUsername;
    m_sPassword         = dlg.m_sPassword;
    m_bAutoDial         = dlg.m_bAutoDial;
    m_sBoundIP          = dlg.m_sBoundIP;
    m_sEncodingCharset  = dlg.m_sEncodingCharset;
    m_sEncodingFriendly = dlg.m_sEncodingFriendly;
    m_bMime             = dlg.m_bMime;
    m_bHTML             = dlg.m_bHTML;
    m_Priority          = static_cast<CPJNSMTPMessage::PRIORITY>(dlg.m_Priority);
    m_ConnectionType    = static_cast<CPJNSMTPConnection::ConnectionType>(dlg.m_ConnectionType);
    m_SSLProtocol       = static_cast<CPJNSMTPConnection::SSLProtocol>(dlg.m_SSLProtocol);
    m_bDSN              = dlg.m_bDSN;
    m_bDSNSuccess       = dlg.m_bDSNSuccess;
    m_bDSNFailure       = dlg.m_bDSNFailure;
    m_bDSNDelay         = dlg.m_bDSNDelay;
    m_bDSNHeaders       = dlg.m_bDSNHeaders;
    m_bMDN              = dlg.m_bMDN;
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
  {
    pMessage->m_From = m_sAddress;
  }
  else 
  {
    CPJNSMTPAddress address(m_sName, m_sAddress);
    pMessage->m_From = address;
  }

#ifndef _DEBUG
  pMessage->m_sXMailer.Empty(); //Release builds do not include a XMailer header
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
    pMessage->m_MessageDispositionEmailAddresses.Add(m_sAddress);

#ifdef _DEBUG
  //Add one custom header (for test purpose)
  pMessage->AddCustomHeader(_T("X-Program: CSTMPMessageTester"));

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

        CString sHost;
        BOOL bSend = TRUE;
        if (m_bDNSLookup)
        {
          if (pMessage->m_To.GetSize() == 0)
          {
            AfxMessageBox(_T("At least one receipient must be specified to use the DNS lookup option"));
            bSend = FALSE;
          }
          else
          {
            CString sAddress(pMessage->m_To.ElementAt(0).m_sEmailAddress);
            int nAmpersand = sAddress.Find(_T("@"));
            if (nAmpersand == -1)
            {
              CString sMsg;
              sMsg.Format(_T("Unable to determine the domain for the email address \"%s\""), sAddress.operator LPCTSTR());
              AfxMessageBox(sMsg);
              bSend = FALSE;
            }
            else
            {
              //We just pick the first MX record found, other implementations could ask the user
              //or automatically pick the lowest priority record
              CString sDomain(sAddress.Right(sAddress.GetLength() - nAmpersand - 1));
              CStringArray servers;
              CWordArray priorities;
              if (!m_Connection.MXLookup(sDomain, servers, priorities))
              {
                CString sMsg;
                sMsg.Format(_T("Unable to perform a DNS MX lookup for the domain \"%s\", Error Code:%u"), sDomain.operator LPCTSTR(), GetLastError());
                AfxMessageBox(sMsg);
                bSend = FALSE;
              }
              else
                sHost = servers.GetAt(0);
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
          m_Connection.Connect(sHost, m_Auth, m_sUsername, m_sPassword, m_nPort, m_ConnectionType, pMessage->m_dwDSN);
          m_Connection.SendMessage(*pMessage);

          //Auto disconnect from the internet
          if (m_bAutoDial)
            m_Connection.CloseInternetConnection();
        }
      }
      catch(CPJNSMTPException* pEx)
      {
        //Display the error
        CString sMsg;
        if (HRESULT_FACILITY(pEx->m_hr) == FACILITY_ITF)
          sMsg.Format(_T("An error occured sending the message, Error:%08X\nDescription:%s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR());
        else
          sMsg.Format(_T("An error occured sending the message, Error:%08X\nDescription:%s, Response:%s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR(), pEx->m_sLastResponse.operator LPCTSTR());
        AfxMessageBox(sMsg, MB_ICONSTOP);
  
        pEx->Delete();
        bCloseGracefully = FALSE;
      }

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
      catch (CPJNSMTPException* pEx)
      {
        pEx->Delete();
      }
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
    CString sNewFile = dlg.GetPathName();
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
}

void CPJNSMTPAppDlg::OnSendFromDisk() 
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
          if (m_bDirectly)
          {
            CString sENVID;
            connection.SendMessage(dlg.GetPathName(), Recipients, from, sENVID);
          }
          else
          {
            //The progress dialog instance
            CSendingDlg dlg2;

            //Let the connection class know about the progress dialog
            connection.m_pProgressDlg = &dlg2;

            //Fill in the structure we need to pass to the worker function
            ThreadSendData tsd;
            tsd.pConnection = &connection;
            tsd.pProgressDlg = &dlg2;
            tsd.pRecipients = &Recipients;
            tsd.pFrom = &from;
            tsd.dwMainThreadID = GetCurrentThreadId();
            tsd.sFilename = dlg.GetPathName();

            //Create the worker thread which will do the actual sending
            CWinThread* pWorkerThread = AfxBeginThread(_SendingWorkerThread, &tsd, THREAD_PRIORITY_NORMAL, CREATE_SUSPENDED);
            pWorkerThread->m_bAutoDelete = FALSE;
            pWorkerThread->ResumeThread();

            //Bring up the progress dialog
            dlg2.DoModal();

            //Wait for the thread to close down
            WaitForSingleObject(pWorkerThread->m_hThread, INFINITE);
            delete pWorkerThread;

            if (!tsd.bSentOK)
            {
              CString sMsg;
              sMsg.Format(_T("An error occured sending the message, Error:%08X\nDescription:%s"), tsd.ExceptionHr, tsd.sExceptionMessage.operator LPCTSTR());
              AfxMessageBox(sMsg, MB_ICONSTOP);
            }
          }
        }
        catch(CPJNSMTPException* pEx)
        {
          //Display the error
          CString sMsg;
          sMsg.Format(_T("An error occured sending the message, Error:%08X\nDescription: %s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR());
          AfxMessageBox(sMsg, MB_ICONSTOP);
  
          pEx->Delete();
        }

        //Auto disconnect from the internet
        if (m_bAutoDial)
          connection.CloseInternetConnection();
      }	
    }
  }
}

void CPJNSMTPAppDlg::OnSave() 
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
      catch(CPJNSMTPException* pEx)
      {
        CString sMsg;
        sMsg.Format(_T("Failed to save message to disk!, Error:%08X, Description:%s"), pEx->m_hr, pEx->GetErrorMessage().operator LPCTSTR());
        AfxMessageBox(sMsg, MB_ICONSTOP);

        pEx->Delete();
      }

      //Tidy up the heap memory
      delete pMessage;
    }	
  }
}

UINT CPJNSMTPAppDlg::_SendingWorkerThread(LPVOID lpParam)
{
  //Validate our parameters
  ThreadSendData* pTSD = static_cast<ThreadSendData*>(lpParam);
  AFXASSUME(pTSD);
  AFXASSUME(pTSD->pConnection);
  AFXASSUME(pTSD->pRecipients);
  AFXASSUME(pTSD->pFrom);

  AttachThreadInput(pTSD->dwMainThreadID, GetCurrentThreadId(), TRUE);

  //Wait until the dialog has been created before we continue
  WaitForSingleObject(pTSD->pProgressDlg->m_evtInitialized, INFINITE);

  //Do the actual sending
  try
  {
    CString sENVID;
    pTSD->pConnection->SendMessage(pTSD->sFilename, *pTSD->pRecipients, *pTSD->pFrom, sENVID);
    pTSD->bSentOK = !pTSD->pProgressDlg->AbortFlag();
  }
  catch(CPJNSMTPException* pEx)
  {
    pTSD->ExceptionHr = pEx->m_hr;
    pTSD->sExceptionMessage = pEx->GetErrorMessage();
    pEx->Delete();
    pTSD->bSentOK = FALSE;
  }

  pTSD->pProgressDlg->SetSafeToClose();
  pTSD->pProgressDlg->PostMessage(WM_CLOSE);

  return 0L;
}
