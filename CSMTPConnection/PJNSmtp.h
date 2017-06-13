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


#ifdef CPJNSMTP_MFC_EXTENSIONS

#ifndef __AFXTEMPL_H__
#pragma message("To avoid this message, put afxtempl.h in your pre compiled header (usually stdafx.h)")
#include <afxtempl.h>
#endif //#ifndef __AFXTEMPL_H__

typedef CString CPJNSMTPString;
typedef CStringA CPJNSMTPAsciiString;
typedef CStringArray CPJNSMTPStringArray;
typedef CWordArray CPJNSMTPWordArray;

#else

#ifndef _EXCEPTION_
#pragma message("To avoid this message, please put exception in your pre compiled header (normally stdafx.h)")
#include <exception>
#endif //#ifndef _EXCEPTION_

#ifndef _STRING_
#pragma message("To avoid this message, please put string in your pre compiled header (normally stdafx.h)")
#include <string>
#endif //#ifndef _STRING_

#ifndef _SSTREAM_
#pragma message("To avoid this message, please put sstream in your pre compiled header (normally stdafx.h)")
#include <sstream>
#endif //#ifndef _SSTREAM_

#ifndef _VECTOR_
#pragma message("To avoid this message, please put vector in your pre compiled header (normally stdafx.h)")
#include <vector>
#endif //#ifndef _VECTOR_

#ifndef _ALGORITHM_
#pragma message("To avoid this message, please put algorithm in your pre compiled header (normally stdafx.h)")
#include <algorithm>
#endif //#ifndef _ALGORITHM_

#ifdef _UNICODE
typedef std::wstring CPJNSMTPString;
#else
typedef std::string CPJNSMTPString;
#endif //#ifdef _UNICODE
typedef std::string CPJNSMTPAsciiString;
typedef std::vector<CPJNSMTPString> CPJNSMTPStringArray;
typedef std::vector<WORD> CPJNSMTPWordArray;

#endif //#ifndef CPJNSMTP_MFC_EXTENSIONS


/////////////////////////////// Classes ///////////////////////////////////////

///////////// Class which makes using the ATL base64 encoding / decoding functions somewhat easier

class PJNSMTP_EXT_CLASS CPJNSMPTBase64Encode
{
public:
//Constructors / Destructors
  CPJNSMPTBase64Encode();

//Methods
  void	Encode(_In_reads_bytes_(nSize) const BYTE* pbyData, _In_ int nSize, _In_ DWORD dwFlags);
  void	Decode(_In_reads_bytes_(nSize) LPCSTR pData, _In_ int nSize);
  void	Encode(_In_z_ LPCSTR pszMessage, _In_ DWORD dwFlags);
  void	Decode(_In_z_ LPCSTR sMessage);

  LPSTR Result() const { return m_Buf.m_pData; };
  int	  ResultSize() const { return m_nSize; };

protected:
//Member variables
  ATL::CHeapPtr<char> m_Buf;
  int                 m_nSize;
};


///////////// Class which makes using the ATL QP encoding functions somewhat easier

class PJNSMTP_EXT_CLASS CPJNSMPTQPEncode
{
public:
//Constructors / Destructors
  CPJNSMPTQPEncode();

//Methods
  void	Encode(_In_reads_bytes_(nSize) const BYTE* pbyData, _In_ int nSize, _In_ DWORD dwFlags);
  void	Encode(_In_z_ LPCSTR pszMessage, _In_ DWORD dwFlags);
  LPSTR Result() const { return m_Buf.m_pData; };
  int	  ResultSize() const { return m_nSize; };

protected:
//Member variables
  ATL::CHeapPtr<char> m_Buf;
  int                 m_nSize;
};


///////////// Class which makes using the ATL Q encoding functions somewhat easier

class PJNSMTP_EXT_CLASS CPJNSMPTQEncode
{
public:
//Constructors / Destructors
  CPJNSMPTQEncode();

//Methods
  void	Encode(_In_reads_bytes_(nSize) const BYTE* pbyData, _In_ int nSize, _In_z_ LPCSTR szCharset);
  void	Encode(_In_z_ LPCSTR pszMessage, _In_z_ LPCSTR szCharset);
  LPSTR Result() const { return m_Buf.m_pData; };
  int	  ResultSize() const { return m_nSize; };

protected:
//Member variables
  ATL::CHeapPtr<char> m_Buf;
  int                 m_nSize;
};


///////////// Exception class /////////////////////////////////////////////////

#ifdef CPJNSMTP_MFC_EXTENSIONS
class PJNSMTP_EXT_CLASS CPJNSMTPException : public CException
#else
class PJNSMTP_EXT_CLASS CPJNSMTPException : public std::exception
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
{
public:
//Constructors / Destructors
  CPJNSMTPException(_In_ HRESULT hr, _In_ const CPJNSMTPString& sLastResponse);
  CPJNSMTPException(_In_ DWORD dwError, _In_ DWORD dwFacility, _In_ const CPJNSMTPString& sLastResponse);

//Methods
#if _MSC_VER >= 1700
  virtual BOOL GetErrorMessage(_Out_writes_z_(nMaxError) LPTSTR lpszError, _In_ UINT nMaxError, _Out_opt_ PUINT pnHelpContext = NULL);
#else
  virtual BOOL GetErrorMessage(__out_ecount_z(nMaxError) LPTSTR lpszError, __in UINT nMaxError, __out_opt PUINT pnHelpContext = NULL);
#endif //#if _MSC_VER >= 1700

#ifdef CPJNSMTP_MFC_EXTENSIONS
#ifdef _DEBUG
  virtual void Dump(CDumpContext& dc) const;
#endif //#ifdef _DEBUG

  CPJNSMTPString GetErrorMessage();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

//Members variables
  HRESULT        m_hr;
  CPJNSMTPString m_sLastResponse;
};


////// Encapsulation of an SMTP email address /////////////////////////////////

class PJNSMTP_EXT_CLASS CPJNSMTPAddress
{
public: 
//Constructors / Destructors
  CPJNSMTPAddress();
  CPJNSMTPAddress(_In_ const CPJNSMTPAddress& address);
  CPJNSMTPAddress(_In_z_ LPCTSTR pszAddress);
  CPJNSMTPAddress(_In_z_ LPCTSTR pszFriendly, _In_z_ LPCTSTR pszAddress);
  CPJNSMTPAddress& operator=(_In_ const CPJNSMTPAddress& r);

//Methods
  CPJNSMTPAsciiString GetRegularFormat(_In_ BOOL bEncode, _In_ const CPJNSMTPString& sCharset) const;

//Data members
  CPJNSMTPString m_sFriendlyName; //Would set it to contain something like "PJ Naughter"
  CPJNSMTPString m_sEmailAddress; //Would set it to contains something like "pjna@naughter.com"
  
protected:
//Methods
#ifndef CPJNSMTP_MFC_EXTENSIONS
  static void StringTrim(_Inout_ CPJNSMTPString& sString);
  static void StringReplace(_Inout_ std::string& sString, _In_z_ LPCSTR pszToFind, _In_z_ LPCSTR pszToReplace);
  static void StringReplace(_Inout_ std::wstring& sString, _In_z_ LPCWSTR pszToFind, _In_z_ LPCWSTR pszToReplace);
#endif //#ifndef CPJNSMTP_MFC_EXTENSIONS
  static CPJNSMTPString RemoveQuotes(_In_ const CPJNSMTPString& sValue);

  friend class CPJNSMTPException;
  friend class CPJNSMTPMessage;
  friend class CPJNSMTPBodyPart;
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
  CPJNSMTPBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart);
  CPJNSMTPBodyPart& operator=(_In_ const CPJNSMTPBodyPart& bodyPart);
  virtual ~CPJNSMTPBodyPart();

//Accessors / Mutators
  virtual BOOL                      SetFilename(_In_ const CPJNSMTPString& sFilename);
  virtual CPJNSMTPString            GetFilename() const { return m_sFilename; };
  virtual void                      SetRawBody(_In_z_ LPCSTR pszRawBody);
  virtual LPCSTR                    GetRawBody();
  virtual void                      SetText(_In_ const CPJNSMTPString& sText);
  virtual CPJNSMTPString            GetText() const { return m_sText; };
  virtual void                      SetTitle(_In_z_ LPCTSTR pszTitle) { m_sTitle = pszTitle; };
  virtual CPJNSMTPString            GetTitle() const { return m_sTitle; };
  virtual void                      SetContentType(_In_z_ LPCTSTR pszContentType) { m_sContentType = pszContentType; };
  virtual CPJNSMTPString            GetContentType() const { return m_sContentType; };
  virtual void                      SetCharset(_In_z_ LPCTSTR pszCharset) { m_sCharset = pszCharset; };
  virtual CPJNSMTPString            GetCharset() const { return m_sCharset; };
  virtual void                      SetContentID(_In_z_ LPCTSTR pszContentID);
  virtual CPJNSMTPString            GetContentID() const;
  virtual void                      SetContentLocation(_In_z_ LPCTSTR pszContentLocation);
  virtual CPJNSMTPString            GetContentLocation() const;
  virtual void                      SetAttachment(_In_ BOOL bAttachment);
  virtual BOOL                      GetAttachment() const { return m_bAttachment; };
  virtual void                      SetContentTransferEncoding(_In_ CONTENT_TRANSFER_ENCODING encoding) { m_ContentTransferEncoding = encoding; };
  virtual CONTENT_TRANSFER_ENCODING GetContentTransferEncoding() const { return m_ContentTransferEncoding; };
  virtual void                      SetBoundary(_In_z_ LPCTSTR pszBoundary) { m_sBoundary = pszBoundary; };
  virtual CPJNSMTPString            GetBoundary() const { return m_sBoundary; };

//Misc methods
  virtual CPJNSMTPAsciiString       GetHeader(_In_ const CPJNSMTPString& sRootCharset);
  virtual CPJNSMTPAsciiString       GetBody(_In_ BOOL bDoSingleDotFix);
  virtual CPJNSMTPAsciiString       GetFooter();
  virtual CPJNSMTPBodyPart*         FindFirstBodyPart(_In_ const CPJNSMTPString& sContentType);

//Child Body part methods
  virtual INT_PTR                   GetNumberOfChildBodyParts() const;
  virtual INT_PTR                   AddChildBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart);
  virtual INT_PTR                   AddChildBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart);
  virtual void                      RemoveChildBodyPart(_In_ INT_PTR nIndex);
  virtual void                      InsertChildBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart);
  virtual void                      InsertChildBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart);
  virtual CPJNSMTPBodyPart*         GetChildBodyPart(_In_ INT_PTR nIndex);
  virtual CPJNSMTPBodyPart*         GetParentBodyPart();

//Static methods
  static CPJNSMTPAsciiString        ConvertToUTF8(_In_z_ LPCTSTR pszIn);
  static CPJNSMTPString             CreateGUID();
  static CPJNSMTPAsciiString        HeaderEncode(_In_ const CPJNSMTPString& sText, _In_ const CPJNSMTPString& sCharset);
  static CPJNSMTPString             HeaderEncodeT(_In_ const CPJNSMTPString& sText, _In_ const CPJNSMTPString& sCharset);
  static CPJNSMTPAsciiString        FoldSubjectHeader(_In_ const CPJNSMTPString& sSubject, _In_ const CPJNSMTPString& sCharset);

protected:
//Member variables
  CPJNSMTPString                                m_sFilename;           //The file you want to attach
  CPJNSMTPString                                m_sTitle;              //What is it to be know as when emailed
  CPJNSMTPString                                m_sContentType;        //The mime content type for this body part
  CPJNSMTPString                                m_sCharset;            //The charset for this body part
  CPJNSMTPString                                m_sContentID;          //The uniqiue ID for this body part (allows other body parts to refer to us via a CID URL)
  CPJNSMTPString                                m_sContentLocation;    //The relative URL for this body part (allows other body parts to refer to us via a relative URL)
  CPJNSMTPString                                m_sText;               //If using strings rather than file, then this is it!
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CArray<CPJNSMTPBodyPart*, CPJNSMTPBodyPart*&> m_ChildBodyParts;      //Child body parts for this body part
#else
  std::vector<CPJNSMTPBodyPart*>                m_ChildBodyParts;      //Child body parts for this body part
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPBodyPart*                             m_pParentBodyPart;     //The parent body part for this body part
  CPJNSMTPString                                m_sBoundary;           //String which is used as the body separator for all child mime parts
  BOOL                                          m_bAttachment;         //Is this body part to be considered an attachment
  LPCSTR                                        m_pszRawBody;          //The raw representation of the body part body. This takes precendence over the other forms of a body part body
  CONTENT_TRANSFER_ENCODING                     m_ContentTransferEncoding; //How should this body part be encoded

//Methods
  static void FixSingleDot(_Inout_ CPJNSMTPAsciiString& sBody);

  friend class CPJNSMTPMessage;
  friend class CPJNSMTPConnection;
};


class PJNSMTP_EXT_CLASS CPJNSMTPConnection;


/////// Encapsulation of an SMTP message //////////////////////////////////////

class PJNSMTP_EXT_CLASS CPJNSMTPMessage
{
public:
//Typedefs
#ifdef CPJNSMTP_MFC_EXTENSIONS
  typedef CArray<CPJNSMTPAddress, CPJNSMTPAddress&> CAddressArray;
#else
  typedef std::vector<CPJNSMTPAddress> CAddressArray;
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

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
  CPJNSMTPMessage(_In_ const CPJNSMTPMessage& message);
  CPJNSMTPMessage& operator=(_In_ const CPJNSMTPMessage& message);

//Recipient support
  CAddressArray m_To;
  CAddressArray m_CC;
  CAddressArray m_BCC;
  static INT_PTR              ParseMultipleRecipients(_In_z_ LPCTSTR pszRecipients, _Inout_ CAddressArray& recipients);

//Body Part support
  virtual INT_PTR             GetNumberOfBodyParts() const;
  virtual INT_PTR             AddBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart);
  virtual INT_PTR             AddBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart);
  virtual void                InsertBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart);
  virtual void                InsertBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart);
  virtual void                RemoveBodyPart(_In_ INT_PTR nIndex);
  virtual CPJNSMTPBodyPart*   GetBodyPart(_In_ INT_PTR nIndex);
  virtual int                 AddMultipleAttachments(_In_z_ LPCTSTR pszAttachments);

//Misc methods
  virtual CPJNSMTPAsciiString GetHeader();
  virtual void                AddTextBody(_In_z_ LPCTSTR pszBody, _In_z_ LPCTSTR pszRootMIMEType = _T("multipart/mixed"));
  virtual CPJNSMTPString      GetTextBody();
  virtual void                AddHTMLBody(_In_z_ LPCTSTR pszBody, _In_z_ LPCTSTR pszRootMIMEType = _T("multipart/mixed"));
  virtual CPJNSMTPString      GetHTMLBody();
  virtual void                SetCharset(_In_z_ LPCTSTR pszCharset);
  virtual CPJNSMTPString      GetCharset() const;
  virtual void                SetMime(BOOL bMime);
  virtual BOOL                GetMime() const { return m_bMime; };
  virtual void                SaveToDisk(_In_z_ LPCTSTR pszFilename);
                                
//Data Members
  CPJNSMTPAddress             m_From;
  CPJNSMTPString              m_sSubject;
  CPJNSMTPString              m_sXMailer;
  CAddressArray               m_ReplyTo;
  CPJNSMTPBodyPart            m_RootPart;
  PRIORITY                    m_Priority;
  DSN_RETURN_TYPE             m_DSNReturnType;
  DWORD                       m_dwDSN;                            //To be filled in with the PJNSMTP_DSN_... flags
  CPJNSMTPAsciiString         m_sENVID;                           //The "Envelope ID" to use for requesting DSN's. If you leave this empty when you are sending the message
                                                                  //then one will be generated for you based on a GUID and you can examine/store this value after the 
                                                                  //message was sent
  CPJNSMTPStringArray         m_MessageDispositionEmailAddresses; //Set this value to an email address if you want to request Message Disposition notifications aka read receipts
  BOOL                        m_bAddressHeaderEncoding;           //Should Address headers be encoded
  CPJNSMTPStringArray         m_CustomHeaders;                    //Any custom headers you want to add

protected:
//Methods
  virtual void                WriteToDisk(ATL::CAtlFile& file, CPJNSMTPBodyPart* pBodyPart, BOOL bRoot, const CPJNSMTPString& sRootCharset);
  virtual CPJNSMTPAsciiString FormDateHeader();

//Member variables
  BOOL                        m_bMime;

  friend class CPJNSMTPConnection;
};


//////// The main class which encapsulates the SMTP connection ////////////////

#ifndef CPJNSMTP_NONTLM
class PJNSMTP_EXT_CLASS CPJNSMTPConnection : public CNTLMClientAuth
#ifndef CPJNSMTP_NOSSL
                                           , protected SSLWrappers::CSocket
#endif //#ifndef CPJNSMTP_NOSSL
#else
class PJNSMTP_EXT_CLASS CPJNSMTPConnection
#ifndef CPJNSMTP_NOSSL
                                           : protected SSLWrappers::CSocket
#endif //#ifndef CPJNSMTP_NOSSL
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
    AUTH_AUTO     = 5, //Try to auto negotiate the authentication protocol to use, the order used will be as decided by the "ChooseAuthenticationMethod" virtual method
    AUTH_XOAUTH2  = 6  //XOAUTH2 authentication. For more information please see https://developers.google.com/gmail/xoauth2_protocol
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
  BOOL           IsConnected() const { return m_bConnected; };
  CPJNSMTPString GetLastCommandResponse() const { return m_sLastCommandResponse; };
  int            GetLastCommandResponseCode() const { return m_nLastCommandResponseCode; };
  DWORD          GetTimeout() const { return m_dwTimeout; };
  void           SetTimeout(_In_ DWORD dwTimeout) { m_dwTimeout = dwTimeout; };
  void           SetHeloHostname(_In_z_ LPCTSTR pszHostname);
  CPJNSMTPString GetHeloHostName() const { return m_sHeloHostname; };
  void           SetSSLProtocol(_In_ SSLProtocol protocol) { m_SSLProtocol = protocol; }
  SSLProtocol    GetSSLProtocol() const { return m_SSLProtocol; }
  void           SetBindAddress(_In_z_ LPCTSTR pszBindAddress) { m_sBindAddress = pszBindAddress; };
  CPJNSMTPString GetBindAddress() const { return m_sBindAddress; };
  BOOL           ImplementsDSN() const { return m_bCanDoDSN; }; //Note you should only use this method after calling Connect where the variable actually gets set
  BOOL           ImplementsSTARTTLS() const { return m_bCanDoSTARTTLS; }; //Note you should only use this method after calling Connect where the variable actually gets set
  void           SetInitialClientResponse(_In_ BOOL bDoInitialClientResponse) { m_bDoInitialClientResponse = bDoInitialClientResponse; };
  BOOL           GetInitialClientResponse() const { return m_bDoInitialClientResponse; };

//"Wininet" Connectivity methods
  ConnectToInternetResult ConnectToInternet();
  BOOL CloseInternetConnection();

//MX Lookup support
  BOOL MXLookup(_In_z_ LPCTSTR lpszHostDomain, _Inout_ CPJNSMTPStringArray& arrHosts, _Inout_ CPJNSMTPWordArray& arrPreferences, _In_ WORD fOptions = DNS_QUERY_STANDARD, _In_opt_ PIP4_ARRAY aipServers = NULL);

//Static methods
  static void ThrowPJNSMTPException(_In_ DWORD dwError = 0, _In_ DWORD Facility = FACILITY_WIN32, _In_ const CPJNSMTPString& sLastResponse = _T(""));
  static void ThrowPJNSMTPException(_In_ HRESULT hr, _In_ const CPJNSMTPString& sLastResponse = _T(""));
  static CPJNSMTPAsciiString CreateNEWENVID();

//Virtual Methods
  virtual void SendMessage(_Inout_ CPJNSMTPMessage& message);
  virtual void SendMessage(_In_z_ LPCTSTR pszMessageOnFile, _Inout_ CPJNSMTPMessage::CAddressArray& Recipients, _In_ const CPJNSMTPAddress& From, _Inout_ CPJNSMTPAsciiString& sENVID, _In_ DWORD dwSendBufferSize = 4096, _In_ DWORD dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED, _In_ CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType = CPJNSMTPMessage::HeadersOnly);
  virtual void SendMessage(_In_reads_bytes_opt_(dwMessageSize) const BYTE* pMessage, _In_ DWORD dwMessageSize, _Inout_ CPJNSMTPMessage::CAddressArray& Recipients, _In_ const CPJNSMTPAddress& From, _Inout_ CPJNSMTPAsciiString& sENVID, _In_ DWORD dwSendBufferSize = 4096, _In_ DWORD dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED, _In_ CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType = CPJNSMTPMessage::HeadersOnly);
  virtual void Connect(_In_z_ LPCTSTR pszHostName, _In_ AuthenticationMethod am = AUTH_NONE, _In_opt_z_ LPCTSTR pszUsername = NULL, _In_opt_z_ LPCTSTR pszPassword = NULL, _In_ int nPort = 25, _In_ ConnectionType connectionType = PlainText, _In_ DWORD dwDSN = CPJNSMTPMessage::DSN_NOT_SPECIFIED);
  void         Disconnect(_In_ BOOL bGracefully = TRUE);
  virtual BOOL OnSendProgress(_In_ DWORD dwCurrentBytes, _In_ DWORD dwTotalBytes);

protected:
//Methods
#ifndef CPJNSMTP_NOSSL
  virtual void DoSTARTTLS();
  virtual void CreateSSLSocket();
  virtual BOOL ReadSSLResponse(_Inout_ CPJNSMTPAsciiString& sResponse);
#endif //#ifndef CPJNSMTP_NOSSL
  virtual void AuthCramMD5(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszPassword);
  virtual void ConnectESMTP(_In_z_ LPCTSTR pszLocalName, 
                            _When_(am == AUTH_CRAM_MD5, _In_z_)
                            _When_(am == AUTH_LOGIN, _In_z_)
                            _When_(am == AUTH_PLAIN, _In_z_)
                            _When_(am == AUTH_XOAUTH2, _In_z_) 
                            _When_((am != AUTH_CRAM_MD5) && (am != AUTH_LOGIN) && (am != AUTH_PLAIN) && (am != AUTH_XOAUTH2), _In_opt_z_) 
                            LPCTSTR pszUsername, 
                            _When_(am == AUTH_CRAM_MD5, _In_z_)
                            _When_(am == AUTH_LOGIN, _In_z_)
                            _When_(am == AUTH_PLAIN, _In_z_)
                            _When_(am == AUTH_XOAUTH2, _In_z_) 
                            _When_((am != AUTH_CRAM_MD5) && (am != AUTH_LOGIN) && (am != AUTH_PLAIN) && (am != AUTH_XOAUTH2), _In_opt_z_) 
                            LPCTSTR pszPassword, _In_ AuthenticationMethod am);
  virtual void ConnectSMTP(_In_z_ LPCTSTR pszLocalName);
  virtual AuthenticationMethod ChooseAuthenticationMethod(_In_z_ LPCTSTR pszAuthMethods);
  virtual void AuthLogin(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszPassword);
  virtual void AuthPlain(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszPassword);
  virtual void AuthXOAUTH2(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszAccessToken);
  virtual CPJNSMTPAsciiString FormMailFromCommand(_In_z_ LPCTSTR pszEmailAddress, _In_ DWORD dwDSN, _In_ CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType, _Inout_ CPJNSMTPAsciiString& sENVID);
  virtual void SendRCPTForRecipient(_In_ DWORD dwDSN, _Inout_ CPJNSMTPAddress& recipient);
  virtual void SendBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart, _In_ BOOL bRoot, _In_z_ LPCTSTR pszRootCharset);
  virtual BOOL ReadCommandResponse(_In_ int nExpectedCode);
  virtual BOOL ReadCommandResponse(_In_ int nExpectedCode1, _In_ int nExpectedCode2);
  virtual BOOL ReadResponse(_Inout_ CPJNSMTPAsciiString& sResponse);
  virtual BOOL ReadPlainTextResponse(_Inout_ CPJNSMTPAsciiString& sResponse);
#ifndef CPJNSMTP_NONTLM
  virtual SECURITY_STATUS NTLMAuthPhase1(_In_reads_bytes_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf);
  virtual SECURITY_STATUS NTLMAuthPhase2(_Inout_updates_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf, _Out_ DWORD* pcbRead);
  virtual SECURITY_STATUS NTLMAuthPhase3(_In_reads_bytes_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf);
#endif //#ifndef CPJNSMTP_NONTLM
  virtual void _Send(const void *pBuffer, int nBuf);
  virtual BOOL DoEHLO(_In_ AuthenticationMethod am, _In_ ConnectionType connectionType, _In_ DWORD dwDSN);
  virtual void SendEHLO(_In_z_ LPCTSTR pszLocalName, _Inout_ CPJNSMTPString& sLastResponse);

//Member variables
#ifndef CPJNSMTP_NOSSL
  SSLWrappers::CCachedCredentials m_SSLCredentials;               //SSL credentials
#endif //#ifndef CPJNSMTP_NOSSL
  ConnectionType                  m_ConnectionType;               //What type of connection are we using
  CWSocket                        m_Socket;                       //The socket connection to the SMTP server (if not using SSL)
  BOOL                            m_bConnected;                   //Are we currently connected to the server 
  CPJNSMTPString                  m_sLastCommandResponse;         //The full last response the server sent us  
  CPJNSMTPString                  m_sHeloHostname;                //The hostname we will use in the HELO command
  DWORD                           m_dwTimeout;                    //The timeout in milliseconds
  int                             m_nLastCommandResponseCode;     //The last numeric SMTP response
  CPJNSMTPString                  m_sBindAddress;                 //The address we will bind to
  SSLProtocol                     m_SSLProtocol;                  //What flavour of SSL should we speak
  BOOL                            m_bCanDoDSN;                    //Is the server capable of "DSN"
  BOOL                            m_bCanDoSTARTTLS;               //Is the server capable of "STARTTLS"
  BOOL                            m_bDoInitialClientResponse;     //Should we use "SASL-IR" for authentication requests
  CPJNSMTPString                  m_sHostName;                    //The host name we are connecting to
};

#endif //#ifndef __PJNSMTP_H__
