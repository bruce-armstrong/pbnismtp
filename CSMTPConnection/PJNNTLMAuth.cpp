/*
Module : PJNNTLMAuth.CPP
Purpose: Implementation for a simple wrapper class for client NTLM authentication via SSPI
Created: PJN / 06-04-2005
History: PJN / 05-09-2005 1. Function pointer to CompleteAuthToken is now constructed at runtime. This means
                          that NTLM client authentication will work correctly on Win9x as used by the class
                          CPJNSMTPConnection and any other class instead of bringing up a "Failed to load
                          due to missing export..." message from the Windows loader. If you want to remove
                          the NTLM support from CPJNSMTPConnection then you can continue to use the CPJNSMTP_NONTLM
                          preprocessor define. Thanks to "Lior" for reporting this problem.
         PJN / 29-09-2005 1. All SDK function pointers used by the class are now implemented using GetProcAddress.
                          Thanks to Emir Kapic for this update.
         PJN / 29-06-2006 1. Combined the functionality of the _PJNNTLMCLIENTAUTH_DATA class into the main 
                          CNTLMClientAuth class.
         PJN / 14-07-2006 1. Updated CNTLMClientAuth to use newer C++ style casts.
                          2. Made the NTLMAuthenticate function "exception" safe, that is derived classes are free to throw 
                          exceptions in their implementations of NTLMAuthPhase1, NTLMAuthPhase2 or NTLMAuthPhase3 and the
                          CNTLMClientAuth class will properly clean up its resources. This is achieved by making the CredHandle
                          and SecHandle values member variables of the class.
                          3. Replaced all calls to ZeroMemory with memset.
                          4. Now the public NTLMAuthenticate function allows you to pass a user name and password pair to 
                          do NTLM authentication using a specified account instead of the credentials of the current user. If 
                          you leave these values as NULL, then you will get the old behaviour which is to use the current
                          user credentials. Thanks to Wouter Demuynck for this very nice addition.
         PJN / 13-10-2006 1. Code now initializes the Domain name value in NTLMAuthenticate to an empty string instead of a NULL
                          string. This avoids NTLM authentication issues when authenticating as a non-domain user with certain
                          mail servers. Thanks to Wouter Demuynck for reporting this issue.
         PJN / 31-05-2008 1. Code now compiles cleanly using Code Analysis (/analyze)
         PJN / 30-09-2012 1. Updated the code to avoid DLL planting security issues when calling LoadLibrary. Thanks to Mat 
                          Berchtold for reporting this issue.
         PJN / 15-11-2014 1. CNTLMClientAuth now takes a static dependency on Secure32.dll instead of using GetProcAddress.
                          2. The CNTLMClientAuth now calls QuerySecurityPackageInfo to determine the buffer size to use 
                          3. The SEC_SUCCESS macro has now been made a method called "SEC_SUCCESS"
                          4. General cleanup of the CNTLMClientAuth::GenClientContext method
                          5. Made the CNTLMClientAuth destructor virtual
         PJN / 16-01-2015 1. Updated copright details
         PJN / 11-08-2016 1. Minor review and tidy up of the CNTLMClientAuth code
                          2. Reworked the class to optionally compile without MFC. By default the class now uses STL 
                          classes and idioms but if you define CNTLMCLIENTAUTH_MFC_EXTENSIONS the class will revert back to the
                          MFC behaviour.
         PJN / 14-11-2016 1. Updated the non MFC code path in CNTLMClientAuth::NTLMAuthenticate to handle a NULL pszUsername
                          parameter being passed. Thanks to Christopher Craft for reporting this issue.

Copyright (c) 2005 - 2016 by PJ Naughter (Web: www.naughter.com, Email: pjna@naughter.com)

All rights reserved.

Copyright / Usage Details:

You are allowed to include the source code in any product (commercial, shareware, freeware or otherwise) 
when your product is released in binary form. You are allowed to modify the source code in any way you want 
except you cannot modify the copyright details at the top of each module. If you want to distribute source 
code with your application, then you are only allowed to distribute versions released by the author. This is 
to maintain a single distribution point for the source code. 

*/


//////////////// Includes /////////////////////////////////////////////////////

#include "stdafx.h"
#include "PJNNTLMAuth.h"
#ifndef __SSPI_H__
#pragma message("To avoid this message, please put Sspi.h in your pre compiled header (usually stdafx.h)")
#include <Sspi.h>
#endif //#ifndef __SSPI_H__


//////////////// Macros / Locals //////////////////////////////////////////////

#pragma comment(lib, "Secur32.lib") //Automatically link in the Secur32 dll

#ifdef CNTLMCLIENTAUTH_MFC_EXTENSIONS
#ifdef _DEBUG
#define new DEBUG_NEW
#endif //#ifdef _DEBUG
#endif //#ifdef CNTLMCLIENTAUTH_MFC_EXTENSIONS


//////////////// Implementation ///////////////////////////////////////////////

CNTLMClientAuth::CNTLMClientAuth() : m_dwBufferSize(0)
{
  //Set our credentials handles to default values
  memset(&m_hCred, 0, sizeof(m_hCred));
  memset(&m_hContext, 0, sizeof(m_hContext));
}

CNTLMClientAuth::~CNTLMClientAuth()
{
  ReleaseHandles();
}

void CNTLMClientAuth::ReleaseHandles()
{
  //Free up the security context if valid
  if (m_hContext.dwLower != 0 || m_hContext.dwUpper != 0) 
  {
    DeleteSecurityContext(&m_hContext);
    memset(&m_hContext, 0, sizeof(m_hContext));
  }

  //Free up the credentials handle if valid
  if (m_hCred.dwLower != 0 || m_hCred.dwUpper != 0) 
  {
    FreeCredentialsHandle(&m_hCred);
    memset(&m_hCred, 0, sizeof(m_hCred));
  }
}

SECURITY_STATUS CNTLMClientAuth::NTLMAuthenticate(_In_opt_z_ LPCTSTR pszUsername, _In_opt_z_ LPCTSTR pszPassword)
{
  //Allow "pszUsername" parameter to be of the format "DomainName\UserName"
  LPCTSTR pszDomain = _T("");
#ifdef CNTLMCLIENTAUTH_MFC_EXTENSIONS
  CString sUsername(pszUsername);
  int nSlashSeparatorOffset = sUsername.Find(_T('\\'));
  CString sDomain;
  CString sUsernameWithoutDomain;
  if (nSlashSeparatorOffset != -1)
  {
    sDomain = sUsername.Left(nSlashSeparatorOffset);
    pszDomain = sDomain.operator LPCTSTR();
    sUsernameWithoutDomain = sUsername.Mid(nSlashSeparatorOffset + 1);
    pszUsername = sUsernameWithoutDomain.operator LPCTSTR();
  }
#else
#ifdef _UNICODE
  std::wstring sUsername(pszUsername != NULL ? pszUsername : L"");
  std::wstring::size_type nSlashSeparatorOffset = sUsername.find(_T('\\'));
  std::wstring sDomain;
  std::wstring sUsernameWithoutDomain;
  if (nSlashSeparatorOffset != std::wstring::npos)
#else
  std::string sUsername(pszUsername != NULL ? pszUsername : "");
  std::string::size_type nSlashSeparatorOffset = sUsername.find(_T('\\'));
  std::string sDomain;
  std::string sUsernameWithoutDomain;
  if (nSlashSeparatorOffset != std::string::npos)
#endif //#ifdef _UNICODE
  {
    sDomain = sUsername.substr(0, nSlashSeparatorOffset);
    pszDomain = sDomain.c_str();
    sUsernameWithoutDomain = sUsername.substr(nSlashSeparatorOffset + 1);
    pszUsername = sUsernameWithoutDomain.c_str();
  }
#endif //#ifdef CNTLMCLIENTAUTH_MFC_EXTENSIONS

  //Release the handles before we try to authenticate (we do this here to ensure any previous calls
  //to NTLMAuthenticate which throw exceptions are cleaned up prior to any new calls to DoNTLMAuthentication)
  ReleaseHandles();

  //Get the buffer size to use for the NTLM SSPI  
  PSecPkgInfo pSecInfo = NULL;
#ifdef _UNICODE
  SECURITY_STATUS ss = QuerySecurityPackageInfo(NTLMSP_NAME, &pSecInfo);
#else
  SECURITY_STATUS ss = QuerySecurityPackageInfo(NTLMSP_NAME_A, &pSecInfo);
#endif //#ifdef _UNICODE
  if (ss != SEC_E_OK)
    return ss;
  m_dwBufferSize = pSecInfo->cbMaxToken;

  //Call the helper function which does all the work
  ss = DoNTLMAuthentication(pszUsername, pszPassword, pszDomain);

  //Now free up the handles now that we are finished the authentication (note it is not critical that this code is
  //called since the various NTLMAuthPhase(*) functions may throw exceptions)
  ReleaseHandles();

  return ss;
}

BOOL CNTLMClientAuth::SEC_SUCCESS(_In_ SECURITY_STATUS ss)
{
  return (ss == SEC_E_OK) || (ss == SEC_I_COMPLETE_AND_CONTINUE) || (ss == SEC_I_COMPLETE_NEEDED) || (ss == SEC_I_CONTINUE_NEEDED);
}

SECURITY_STATUS CNTLMClientAuth::DoNTLMAuthentication(_In_opt_z_ LPCTSTR pszUsername, _In_opt_z_ LPCTSTR pszPassword, _In_opt_z_ LPCTSTR pszDomain)
{
  //Allocate some heap space to contain the in and out buffers for SSPI
  ATL::CHeapPtr<BYTE> inBuf;
  if (!inBuf.Allocate(m_dwBufferSize))
    return SEC_E_INSUFFICIENT_MEMORY;
    
  ATL::CHeapPtr<BYTE> outBuf;
  if (!outBuf.Allocate(m_dwBufferSize))
    return SEC_E_INSUFFICIENT_MEMORY;
  
  DWORD cbMaxMessage = m_dwBufferSize;
  DWORD cbOut = cbMaxMessage;
  BOOL bDone = FALSE;
  SECURITY_STATUS ss = GenClientContext(NULL, 0, outBuf.m_pData, &cbOut, &bDone, pszUsername, pszPassword, pszDomain);
  if (!SEC_SUCCESS(ss))
    return ss;

  ss = NTLMAuthPhase1(outBuf.m_pData, cbOut);
  if (!SEC_SUCCESS(ss))
    return ss;

  DWORD cbIn = 0;
  while (!bDone) 
  {
    ss = NTLMAuthPhase2(inBuf.m_pData, m_dwBufferSize, &cbIn);
    if (!SEC_SUCCESS(ss))
      return ss;

    cbOut = cbMaxMessage;

    ss = GenClientContext(inBuf.m_pData, cbIn, outBuf.m_pData, &cbOut, &bDone, pszUsername, pszPassword, pszDomain);
    if (!SEC_SUCCESS(ss))
      return ss;

    ss = NTLMAuthPhase3(outBuf.m_pData, cbOut);
    if (!SEC_SUCCESS(ss))
      return ss;
  }

  return ss;
}

SECURITY_STATUS CNTLMClientAuth::GenClientContext(_In_reads_bytes_opt_(cbIn) BYTE* pIn, _In_ DWORD cbIn, _Out_writes_bytes_(*pcbOut) BYTE* pOut, _Inout_ DWORD* pcbOut, _Inout_ BOOL* pfDone, _In_opt_z_ LPCTSTR pszUserName, _In_opt_z_ LPCTSTR pszPassword, _In_opt_z_ LPCTSTR pszDomain)
{
  //Validate our parameters
  ATLASSUME(pcbOut != NULL);
  ATLASSUME(pfDone != NULL);

  SECURITY_STATUS ss = SEC_E_OK;
  if (pIn == NULL)  
  {   
    SEC_WINNT_AUTH_IDENTITY authInfo;
    memset(&authInfo, 0, sizeof(authInfo));
    void* pvLogonID = NULL;
    size_t nUserNameLength = 0;
    if (pszUserName != NULL)
      nUserNameLength = _tcslen(pszUserName);
    if (nUserNameLength)
    {
      authInfo.UserLength = static_cast<unsigned long>(nUserNameLength);
      ATLASSUME(pszDomain != NULL);
      authInfo.DomainLength = static_cast<unsigned long>(_tcslen(pszDomain));
      ATLASSUME(pszPassword != NULL);
      authInfo.PasswordLength = static_cast<unsigned long>(_tcslen(pszPassword));
    #ifdef _UNICODE
      authInfo.User = reinterpret_cast<unsigned short*>(const_cast<LPTSTR>(pszUserName));
      authInfo.Domain = reinterpret_cast<unsigned short*>(const_cast<LPTSTR>(pszDomain));
      authInfo.Password = reinterpret_cast<unsigned short*>(const_cast<LPTSTR>(pszPassword));
      authInfo.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;
    #else
      authInfo.User = reinterpret_cast<unsigned char*>(const_cast<LPTSTR>(pszUserName));
      authInfo.Domain = reinterpret_cast<unsigned char*>(const_cast<LPTSTR>(pszDomain));
      authInfo.Password = reinterpret_cast<unsigned char*>(const_cast<LPTSTR>(pszPassword));
      authInfo.Flags = SEC_WINNT_AUTH_IDENTITY_ANSI;
    #endif //#ifdef _UNICODE
      pvLogonID = &authInfo;
    }

    TimeStamp Lifetime;
    ss = AcquireCredentialsHandle(NULL, _T("NTLM"), SECPKG_CRED_OUTBOUND, NULL, pvLogonID, NULL, NULL, &m_hCred, &Lifetime);
    if (ss != SEC_E_OK)
      return ss;
  }

  //Prepare the buffers
  SecBufferDesc OutBuffDesc;
  OutBuffDesc.ulVersion = 0;
  OutBuffDesc.cBuffers  = 1;
  SecBuffer OutSecBuff;
  OutBuffDesc.pBuffers  = &OutSecBuff;
  OutSecBuff.cbBuffer   = *pcbOut;
  OutSecBuff.BufferType = SECBUFFER_TOKEN;
  OutSecBuff.pvBuffer   = pOut;

  //The input buffer is created only if a message has been received from the server
  SecBufferDesc InBuffDesc;
  SecBuffer InSecBuff;
  if (pIn != NULL)   
  {
    InBuffDesc.ulVersion = 0;
    InBuffDesc.cBuffers  = 1;
    InBuffDesc.pBuffers  = &InSecBuff;
    InSecBuff.cbBuffer   = cbIn;
    InSecBuff.BufferType = SECBUFFER_TOKEN;
    InSecBuff.pvBuffer   = pIn;

    ULONG ContextAttributes = 0;
    ss = InitializeSecurityContext(&m_hCred, &m_hContext, NULL, 0, 0, SECURITY_NATIVE_DREP, &InBuffDesc, 0, &m_hContext, &OutBuffDesc, &ContextAttributes, NULL);
  }
  else
  {
    ULONG ContextAttributes = 0;
    ss = InitializeSecurityContext(&m_hCred, NULL, NULL, 0, 0, SECURITY_NATIVE_DREP, NULL, 0, &m_hContext, &OutBuffDesc, &ContextAttributes, NULL);
  }

  //If necessary, complete the token.
  if ((ss == SEC_I_COMPLETE_NEEDED) || (ss == SEC_I_COMPLETE_AND_CONTINUE))  
  {
    ss = CompleteAuthToken(&m_hContext, &OutBuffDesc);
    if (ss != SEC_E_OK)  
      return ss;
  }
  else if ((ss != SEC_I_CONTINUE_NEEDED) && (ss != SEC_E_OK))
    return ss;

  //Finally update the output parameters
  *pcbOut = OutSecBuff.cbBuffer;
  *pfDone = !((ss == SEC_I_CONTINUE_NEEDED) || (ss == SEC_I_COMPLETE_AND_CONTINUE));

  return ss;
}

