/*
Module : PJNSmtp.h
Purpose: Defines the interface for a MFC class encapsulation of the SMTP protocol
Created: PJN / 22-05-1998

Copyright (c) 1998 - 2016 by PJ Naughter (Web: www.naughter.com, Email: pjna@naughter.com)

All rights reserved.

Copyright / Usage Details:

You are allowed to include the source code in any product (commercial, shareware, freeware or otherwise) 
when your product is released in binary form. You are allowed to modify the source code in any way you want 
except you cannot modify the copyright details at the top of each module. If you want to distribute source 
code with your application, then you are only allowed to distribute versions released by the author. This is 
to maintain a single distribution point for the source code. 

Please note that I have been informed that C(PJN)SMTPConnection is being used to develop and send unsolicted bulk mail. 
This was not the intention of the code and the author explicitly forbids use of the code for any software of this kind.

*/


/////////////////////////////// Macros / Defines //////////////////////////////

#pragma once

#ifndef __PJNSMTP_H__
#define __PJNSMTP_H__

#ifndef _WINDNS_INCLUDED_
#pragma message("To avoid this message, please put WinDNS.h in your pre compiled header (usually stdafx.h)")
#include <WinDNS.h>
#endif //#ifndef _WINDNS_INCLUDED_

#include "SocMFC.h" //If you get a compilation error about this missing header file, then you need to download my CWSocket class from http://www.naughter.com/w3mfc.html
#ifndef CPJNSMTP_NOSSL
#include "SSLWrappers.h" //If you get a compilation error about this missing header file, then you need to download my SSLWrappers code from http://www.naughter.com/sslwrappers.html
#endif //#ifndef CPJNSMTP_NOSSL

#ifndef CPJNSMTP_NONTLM
#include "PJNNTLMAuth.h"
#endif //#ifndef CPJNSMTP_NONTLM

#ifndef __ATLFILE_H__
#pragma message("To avoid this message, please put atlfile.h in your pre compiled header (usually stdafx.h)")
#include <atlfile.h>
#endif //#ifndef __ATLFILE_H__

#ifndef PJNSMTP_EXT_CLASS
#define PJNSMTP_EXT_CLASS
#endif //#ifndef PJNSMTP_EXT_CLASS


/////////////////////////////// Classes ///////////////////////////////////////

///////////// Class which makes using the ATL base64 encoding / decoding functions somewhat easier

class PJNSMTP_EXT_CLASS CPJNSMPTBase64Encode
{
public:
//Constructors / Destructors
  CPJNSMPTBase64Encode();
  ~CPJNSMPTBase64Encode();

//Methods
  void	Encode(const BYTE* pbyData, int nSize, DWORD dwFlags);
  void	Decode(LPCSTR pData, int nSize);
  void	Encode(LPCSTR pszMessage, DWORD dwFlags);
  void	Decode(LPCSTR sMessage);

  LPSTR Result() const { return m_pBuf; };
  int	  ResultSize() const { return m_nSize; };

protected:
//Member variables
  char* m_pBuf;
  int   m_nSize;
};


///////////// Class which makes using the ATL QP encoding functions somewhat easier

class PJNSMTP_EXT_CLASS CPJNSMPTQPEncode
{
public:
//Constructors / Destructors
  CPJNSMPTQPEncode();
  ~CPJNSMPTQPEncode();

//Methods
  void	Encode(const BYTE* pbyData, int nSize, DWORD dwFlags);
  void	Encode(LPCSTR pszMessage, DWORD dwFlags);

  LPSTR Result() const { return m_pBuf; };
  int	  ResultSize() const { return m_nSize; };

protected:
//Member variables
  char* m_pBuf;
  int   m_nSize;
};


///////////// Class which makes using the ATL Q encoding functions somewhat easier

class PJNSMTP_EXT_CLASS CPJNSMPTQEncode
{
public:
//Constructors / Destructors
  CPJNSMPTQEncode();
  ~CPJNSMPTQEncode();

//Methods
  void	Encode(const BYTE* pbyData, int nSize, LPCSTR szCharset);
  void	Encode(LPCSTR pszMessage, LPCSTR szCharset);

  LPSTR Result() const { return m_pBuf; };
  int	  ResultSize() const { return m_nSize; };

protected:
//Member variables
  char* m_pBuf;
  int   m_nSize;
};


///////////// Exception class /////////////////////////////////////////////////

class PJNSMTP_EXT_CLASS CPJNSMTPException : public CException
{
public:
//Constructors / Destructors
  CPJNSMTPException(HRESULT hr, const CString& sLastResponse = _T("")); 
  CPJNSMTPException(DWORD dwError = 0, DWORD dwFacility = FACILITY_WIN32, const CString& sLastResponse = _T(""));

//Methods
#ifdef _DEBUG
  virtual void Dump(CDumpContext& dc) const;
#endif //#ifdef _DEBUG

#if _MSC_VER >= 1700
  virtual BOOL GetErrorMessage(_Out_writes_z_(nMaxError) LPTSTR lpszError, _In_ UINT nMaxError,	_Out_opt_ PUINT pnHelpContext = NULL);
#else
  virtual BOOL GetErrorMessage(__out_ecount_z(nMaxError) LPTSTR lpszError, __in UINT nMaxError, __out_opt PUINT pnHelpContext = NULL);
#endif //#if _MSC_VER >= 1700
  CString GetErrorMessage();

//Members variables
  HRESULT m_hr;
  CString m_sLastResponse;

protected:
  DECLARE_DYNAMIC(CPJNSMTPException)
};


////// Encapsulation of an SMTP email address /////////////////////////////////

class PJNSMTP_EXT_CLASS CPJNSMTPAddress
{
public: 
//Constructors / Destructors
  CPJNSMTPAddress();
  CPJNSMTPAddress(const CPJNSMTPAddress& address);
  CPJNSMTPAddress(const CString& sAddress);
  CPJNSMTPAddress(const CString& sFriendly, const CString& sAddress);
  CPJNSMTPAddress& operator=(const CPJNSMTPAddress& r);

//Methods
  CString GetRegularFormat(BOOL bEncode, const CString& sCharset) const;

//Data members
  CString m_sFriendlyName; //Would set it to contain something like "PJ Naughter"
  CString m_sEmailAddress; //Would set it to contains something like "pjna@naughter.com"
  
protected:
//Methods
  static CString RemoveQuotes(const CString& sValue);
};


////// Encapsulatation of an SMTP MIME body part //////////////////////////////

class PJNSMTP_EXT_CLASS CPJNSMTPBodyPart
{
public:
//Enums
  enum CONTENT_TRANSFER_ENCODING 
  { 
    NO_ENCODING  = 0, 
    BASE64_ENCODING  = 1, 
    QP_ENCODING = 2 
  };

//Constructors / Destructors
  CPJNSMTPBodyPart();
  CPJNSMTPBodyPart(const CPJNSMTPBodyPart& bodyPart);
  CPJNSMTPBodyPart& operator=(const CPJNSMTPBodyPart& bodyPart);
  virtual ~CPJNSMTPBodyPart();

//Accessors / Mutators
  BOOL    SetFilename(const CString& sFilename);
  CString GetFilename() const { return m_sFilename; }; 

  void    SetRawBody(LPCSTR pszRawBody);
  LPCSTR  GetRawBody();

  void    SetText(const CString& sText);
  CString GetText() const { return m_sText; };

  void    SetTitle(const CString& sTitle) { m_sTitle = sTitle; };
  CString GetTitle() const { return m_sTitle; };

  void    SetContentType(const CString& sContentType) { m_sContentType = sContentType; };
  CString GetContentType() const { return m_sContentType; };

  void    SetCharset(const CString& sCharset) { m_sCharset = sCharset; };
  CString GetCharset() const { return m_sCharset; };

  void    SetContentID(const CString& sContentID);
  CString GetContentID() const;

  void    SetContentLocation(const CString& sContentLocation);
  CString GetContentLocation() const;
  
  void    SetAttachment(BOOL bAttachment);
  BOOL    GetAttachment() const { return m_bAttachment; };

  void    SetContentTransferEncoding(CONTENT_TRANSFER_ENCODING encoding) { m_ContentTransferEncoding = encoding; };
  CONTENT_TRANSFER_ENCODING GetContentTransferEncoding() const { return m_ContentTransferEncoding; };

  void    SetBoundary(const CString& sBoundary) { m_sBoundary = sBoundary; };
  CString GetBoundary() const { return m_sBoundary; };

//Misc methods
  CStringA          GetHeader(const CString& sRootCharset);
  CStringA          GetBody(BOOL bDoSingleDotFix);
  CStringA          GetFooter();
  CPJNSMTPBodyPart* FindFirstBodyPart(const CString sContentType);

//Child Body part methods
  INT_PTR           GetNumberOfChildBodyParts() const;
  INT_PTR           AddChildBodyPart(CPJNSMTPBodyPart& bodyPart);
  void              RemoveChildBodyPart(INT_PTR nIndex);
  void              InsertChildBodyPart(CPJNSMTPBodyPart& bodyPart);
  CPJNSMTPBodyPart* GetChildBodyPart(INT_PTR nIndex);
  CPJNSMTPBodyPart* GetParentBodyPart();

//Static methods
  static CStringA   ConvertToUTF8(const CString& sIn);
  static CString    CreateGUID();
  static CStringA   HeaderEncode(const CString& sText, const CString& sCharset);
  static CStringA   FoldSubjectHeader(const CString& sSubject, const CString& sCharset);

protected:
//Member variables
  CString                                       m_sFilename;           //The file you want to attach
  CString                                       m_sTitle;              //What is it to be know as when emailed
  CString                                       m_sContentType;        //The mime content type for this body part
  CString                                       m_sCharset;            //The charset for this body part
  CString                                       m_sContentID;          //The uniqiue ID for this body part (allows other body parts to refer to us via a CID URL)
  CString                                       m_sContentLocation;    //The relative URL for this body part (allows other body parts to refer to us via a relative URL)
  CString                                       m_sText;               //If using strings rather than file, then this is it!
  CArray<CPJNSMTPBodyPart*, CPJNSMTPBodyPart*&> m_ChildBodyParts;      //Child body parts for this body part
  CPJNSMTPBodyPart*                             m_pParentBodyPart;     //The parent body part for this body part
  CString                                       m_sBoundary;           //String which is used as the body separator for all child mime parts
  BOOL                                          m_bAttachment;         //Is this body part to be considered an attachment
  LPCSTR                                        m_pszRawBody;          //The raw representation of the body part body. This takes precendence over the other forms of a body part body
  CONTENT_TRANSFER_ENCODING                     m_ContentTransferEncoding; //How should this body part be encoded

//Methods
  static void FixSingleDotA(CStringA& sBody);
  static void FixSingleDotT(CString& sBody);

  friend class CPJNSMTPMessage;
  friend class CPJNSMTPConnection;
};


////////////////// Forward declaration ////////////////////////////////////////

class PJNSMTP_EXT_CLASS CPJNSMTPConnection;


/////// Encapsulation of an SMTP message //////////////////////////////////////

class PJNSMTP_EXT_CLASS CPJNSMTPMessage
{
public:
//Typedefs
  typedef CArray<CPJNSMTPAddress, CPJNSMTPAddress&> CAddressArray;

//Enums
  enum RECIPIENT_TYPE 
  { 
    TO  = 0, 
    CC  = 1, 
    BCC = 2 
  };
  enum PRIORITY 
  { 
    NoPriority     = 0, 
    LowPriority    = 1, 
    NormalPriority = 2, 
    HighPriority   = 3 
  };
  enum DSN_RETURN_TYPE
  {
    HeadersOnly = 0,
    FullEmail = 1
  };
  enum DNS_FLAGS
  {
    DSN_NOT_SPECIFIED = 0x00, //We are not specifying if we should be using a DSN or not  
    DSN_SUCCESS       = 0x01, //A DSN should be sent back for messages which were successfully delivered
    DSN_FAILURE       = 0x02, //A DSN should be sent back for messages which were not successfully delivered
    DSN_DELAY         = 0x04  //A DSN should be sent back for messages which were delayed
  };

//Constructors / Destructors
  CPJNSMTPMessage();
  CPJNSMTPMessage(const CPJNSMTPMessage& message);
  CPJNSMTPMessage& operator=(const CPJNSMTPMessage& message);

//Recipient support
  CAddressArray m_To;
  CAddressArray m_CC;
  CAddressArray m_BCC;
  static INT_PTR       ParseMultipleRecipients(const CString& sRecipients, CAddressArray& recipients);

//Body Part support
  INT_PTR              GetNumberOfBodyParts() const;
  INT_PTR              AddBodyPart(CPJNSMTPBodyPart& bodyPart);
  void                 InsertBodyPart(CPJNSMTPBodyPart& bodyPart);
  void                 RemoveBodyPart(INT_PTR nIndex);
  CPJNSMTPBodyPart*    GetBodyPart(INT_PTR nIndex);
  int                  AddMultipleAttachments(const CString& sAttachments);

//Misc methods
  virtual CStringA     GetHeader();
  void                 AddTextBody(const CString& sBody, LPCTSTR pszRootMIMEType = _T("multipart/mixed"));
  CString              GetTextBody();
  void                 AddHTMLBody(const CString& sBody, LPCTSTR pszRootMIMEType = _T("multipart/mixed"));
  CString              GetHTMLBody();
  void                 AddCustomHeader(const CString& sHeader);
  CString              GetCustomHeader(int nIndex);
  INT_PTR              GetNumberOfCustomHeaders() const;
  void                 RemoveCustomHeader(INT_PTR nIndex);
  void                 SetCharset(const CString& sCharset);
  CString              GetCharset() const;
  void                 SetMime(BOOL bMime);
  BOOL                 GetMime() const { return m_bMime; };
  void                 SaveToDisk(const CString& sFilename);
                                
//Data Members
  CPJNSMTPAddress      m_From;
  CString              m_sSubject;
  CString              m_sXMailer;
  CAddressArray        m_ReplyTo;
  CPJNSMTPBodyPart     m_RootPart;
  PRIORITY             m_Priority;
  DSN_RETURN_TYPE      m_DSNReturnType;
  DWORD                m_dwDSN;                            //To be filled in with the PJNSMTP_DSN_... flags
  CString              m_sENVID;                           //The "Envelope ID" to use for requesting DSN's. If you leave this empty when you are sending the message
                                                           //then one will be generated for you based on a GUID and you can examine/store this value after the 
                                                           //message was sent
  CStringArray         m_MessageDispositionEmailAddresses; //Set this value to an email address if you want to request Message Disposition
                                                           //notifications aka read receipts
  BOOL                 m_bAddressHeaderEncoding;           //Should Address headers be encoded
                                  

protected:
//Methods
  void             WriteToDisk(ATL::CAtlFile& file, CPJNSMTPBodyPart* pBodyPart, BOOL bRoot, const CString& sRootCharset);
  virtual CStringA FormDateHeader();

//Member variables
  CStringArray m_CustomHeaders;
  BOOL         m_bMime;

  friend class CPJNSMTPConnection;
};


//////// The main class which encapsulates the SMTP connection ////////////////

#ifndef CPJNSMTP_NONTLM
class PJNSMTP_EXT_CLASS CPJNSMTPConnection : public CNTLMClientAuth
#else
class PJNSMTP_EXT_CLASS CPJNSMTPConnection
#endif //#ifndef CPJNSMTP_NONTLM
{
public:

//typedefs
  enum AuthenticationMethod
  {
    AUTH_NONE     = 0, //Use no authentication with the server
    AUTH_CRAM_MD5 = 1, //CRAM (Challenge Response Authention Method) MD5 (RFC 2195). A challenge is generated by the server, and the response is the MD5 HMAC of the challenge using the password as the key. The username is also supplied in the challenge response.
    AUTH_LOGIN    = 2, //Username and password are simply base64 encoded responses to the server
    AUTH_PLAIN    = 3, //Username and password are sent in the clear to the server
    AUTH_NTLM     = 4, //Use the MS NTLM authentication protocol
    AUTH_AUTO     = 5  //Try to auto negotiate the authentication protocol to use, the order used will be as decided by the "ChooseAuthenticationMethod" virtual method
  };

  enum ConnectToInternetResult
  {
    CTIR_Failure=0,
    CTIR_ExistingConnection=1,
    CTIR_NewConnection=2,
  };

  enum ConnectionType
  {
    PlainText = 0,
  #ifndef CPJNSMTP_NOSSL
    SSL_TLS = 1,
    STARTTLS = 2,
    AutoUpgradeToSTARTTLS = 3
  #endif //#ifndef CPJNSMTP_NOSSL
  };

  enum SSLProtocol
  {
    SSLv2orv3 = 0,
    TLSv1 = 1,
    TLSv1_1 = 2,
    TLSv1_2 = 3,
    //Enum 4 was DTLSv1 which was removed in v3.12
    SSLv2 = 5,
    SSLv3 = 6,
    OSDefault = 7,
    AnyTLS = 8
  };

//Constructors / Destructors
  CPJNSMTPConnection();
  virtual ~CPJNSMTPConnection();

//Standard Methods
  BOOL        IsConnected() const	{ return m_bConnected; };
  CString     GetLastCommandResponse() const { return m_sLastCommandResponse; };
  int         GetLastCommandResponseCode() const { return m_nLastCommandResponseCode; };
  DWORD       GetTimeout() const { return m_dwTimeout; };
  void        SetTimeout(DWORD dwTimeout) { m_dwTimeout = dwTimeout; };
  void        SetHeloHostname(const CString& sHostname);
  CString     GetHeloHostName() const { return m_sHeloHostname; };
  void        SetSSLProtocol(SSLProtocol protocol) { m_SSLProtocol = protocol; }
  SSLProtocol GetSSLProtocol() const { return m_SSLProtocol; }
  void        SetBindAddress(const CString& sBindAddress) { m_sBindAddress = sBindAddress; };
  CString     GetBindAddress() const { return m_sBindAddress; };
  BOOL        ImplementsDSN() const { return m_bCanDoDSN; }; //Note you should only use this method after calling Connect where the variable actually gets set
  BOOL        ImplementsSTARTTLS() const { return m_bCanDoSTARTTLS; }; //Note you should only use this method after calling Connect where the variable actually gets set

//"Wininet" Connectivity methods
  ConnectToInternetResult ConnectToInternet();
  BOOL CloseInternetConnection();

//MX Lookup support
  BOOL MXLookup(LPCTSTR lpszHostDomain, CStringArray& arrHosts, CWordArray& arrPreferences, WORD fOptions = DNS_QUERY_STANDARD, PIP4_ARRAY aipServers = NULL);

//Static methods
  static void ThrowPJNSMTPException(DWORD dwError = 0, DWORD Facility = FACILITY_WIN32, const CString& sLastResponse = _T(""));
  static void ThrowPJNSMTPException(HRESULT hr, const CString& sLastResponse = _T(""));
  static CString CreateNEWENVID();

//Virtual Methods
  virtual void SendMessage(CPJNSMTPMessage& message);
  virtual void SendMessage(const CString& sMessageOnFile, CPJNSMTPMessage::CAddressArray& Recipients, const CPJNSMTPAddress& From, CString& sENVID, DWORD dwSendBufferSize = 4096, DWORD dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED, CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType = CPJNSMTPMessage::HeadersOnly);
  virtual void SendMessage(BYTE* pMessage, DWORD dwMessageSize, CPJNSMTPMessage::CAddressArray& Recipients, const CPJNSMTPAddress& From, CString& sENVID, DWORD dwSendBufferSize = 4096, DWORD dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED, CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType = CPJNSMTPMessage::HeadersOnly);
  virtual void Connect(LPCTSTR pszHostName, AuthenticationMethod am = AUTH_NONE, LPCTSTR pszUsername = NULL, LPCTSTR pszPassword = NULL, int nPort = 25, ConnectionType connectionType = PlainText, DWORD dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED);
  void         Disconnect(BOOL bGracefully = TRUE);
  virtual BOOL OnSendProgress(DWORD dwCurrentBytes, DWORD dwTotalBytes);

protected:
//Methods
#ifndef CPJNSMTP_NOSSL
  virtual void DoSTARTTLS();
  virtual void CreateSSLSocket();
  virtual BOOL ReadSSLResponse(CStringA& sResponse);
#endif //#ifndef CPJNSMTP_NOSSL
  virtual void AuthCramMD5(LPCTSTR pszUsername, LPCTSTR pszPassword);
  virtual void ConnectESMTP(LPCTSTR pszLocalName, LPCTSTR pszUsername, LPCTSTR pszPassword, AuthenticationMethod am);
  virtual void ConnectSMTP(LPCTSTR pszLocalName);
  virtual AuthenticationMethod ChooseAuthenticationMethod(const CString& sAuthMethods);
  virtual void AuthLogin(LPCTSTR pszUsername, LPCTSTR pszPassword);
  virtual void AuthPlain(LPCTSTR pszUsername, LPCTSTR pszPassword);
  virtual CString FormMailFromCommand(const CString& sEmailAddress, DWORD dwDSN, CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType, CString& sENVID);
  virtual void SendRCPTForRecipient(DWORD dwDSN, CPJNSMTPAddress& recipient);
  virtual void SendBodyPart(CPJNSMTPBodyPart* pBodyPart, BOOL bRoot, const CString& sRootCharset);
  virtual BOOL ReadCommandResponse(int nExpectedCode);
  virtual BOOL ReadCommandResponse(int nExpectedCode1, int nExpectedCode2);
  virtual BOOL ReadResponse(CStringA& sResponse);
  virtual BOOL ReadPlainTextResponse(CStringA& sResponse);
#ifndef CPJNSMTP_NONTLM
  virtual SECURITY_STATUS NTLMAuthPhase1(PBYTE pBuf, DWORD cbBuf);
  virtual SECURITY_STATUS NTLMAuthPhase2(PBYTE pBuf, DWORD cbBuf, DWORD* pcbRead);
  virtual SECURITY_STATUS NTLMAuthPhase3(PBYTE pBuf, DWORD cbBuf);
#endif //#ifndef CPJNSMTP_NONTLM
  virtual void _Send(const void *pBuffer, int nBuf);
  virtual BOOL DoEHLO(AuthenticationMethod am, ConnectionType connectionType, DWORD dwDSN);
  virtual void SendEHLO(LPCTSTR pszLocalName, CString& sLastResponse);

//Member variables
#ifndef CPJNSMTP_NOSSL
  SSLWrappers::CCachedCredentials m_SSLCredentials;               //SSL credentials
  SSLWrappers::CSocket            m_SSL;                          //SSL socket wrapper
#endif //#ifndef CPJNSMTP_NOSSL
  ConnectionType                  m_ConnectionType;               //What type of connection are we using
  CWSocket                        m_Socket;                       //The socket connection to the SMTP server (if not using SSL)
  BOOL                            m_bConnected;                   //Are we currently connected to the server 
  CString                         m_sLastCommandResponse;         //The full last response the server sent us  
  CString                         m_sHeloHostname;                //The hostname we will use in the HELO command
  DWORD                           m_dwTimeout;                    //The timeout in milliseconds
  int                             m_nLastCommandResponseCode;     //The last numeric SMTP response
  CString                         m_sBindAddress;                 //The address we will bind to
  SSLProtocol                     m_SSLProtocol;                  //What flavour of SSL should we speak
  BOOL                            m_bCanDoDSN;                    //Is the server capable of "DSN"
  BOOL                            m_bCanDoSTARTTLS;               //Is the server capale of "STARTTLS"
  CString                         m_sHostName;                    //The host name we are connecting to
};

#endif //#ifndef __PJNSMTP_H__
