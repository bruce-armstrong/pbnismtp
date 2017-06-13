/*
Module : PJNNTLMAuth.H
Purpose: Defines the interface for a simple wrapper class for client NTLM authentication via SSPI
Created: PJN / 07-03-2005

Copyright (c) 2005 - 2016 by PJ Naughter (Web: www.naughter.com, Email: pjna@naughter.com)

All rights reserved.

Copyright / Usage Details:

You are allowed to include the source code in any product (commercial, shareware, freeware or otherwise) 
when your product is released in binary form. You are allowed to modify the source code in any way you want 
except you cannot modify the copyright details at the top of each module. If you want to distribute source 
code with your application, then you are only allowed to distribute versions released by the author. This is 
to maintain a single distribution point for the source code. 

*/


/////////////////////////////// Defines ///////////////////////////////////////

#pragma once

#ifndef __PJNNTLMAUTH_H__
#define __PJNNTLMAUTH_H__

#ifndef SECURITY_WIN32
#define SECURITY_WIN32
#endif //#ifndef SECURITY_WIN32
#include <security.h>

#ifndef PJNNTLMAUTH_EXT_CLASS
#define PJNNTLMAUTH_EXT_CLASS
#endif //#ifndef PJNNTLMAUTH_EXT_CLASS


/////////////////////////////// Classes ///////////////////////////////////////

//A class which performs NTLM client authentication
class PJNNTLMAUTH_EXT_CLASS CNTLMClientAuth
{
public:
//Constructors / Destructors
  CNTLMClientAuth();
  virtual ~CNTLMClientAuth();

//Methods
  SECURITY_STATUS NTLMAuthenticate(_In_opt_z_ LPCTSTR pszUsername = NULL, _In_opt_z_ LPCTSTR pszPassword = NULL);
  DWORD GetBufferSize() const { return m_dwBufferSize; };
  void SetBufferSize(_In_ DWORD dwBufferSize) { m_dwBufferSize = dwBufferSize; };

//Virtual methods
  virtual SECURITY_STATUS NTLMAuthPhase1(_In_reads_bytes_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf) = 0;
  virtual SECURITY_STATUS NTLMAuthPhase2(_Inout_updates_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf, _Out_ DWORD* pcbRead) = 0;
  virtual SECURITY_STATUS NTLMAuthPhase3(_In_reads_bytes_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf) = 0;

protected:
//Misc Methods
  static BOOL SEC_SUCCESS(_In_ SECURITY_STATUS ss);
  SECURITY_STATUS GenClientContext(_In_reads_bytes_opt_(cbIn) BYTE* pIn, _In_ DWORD cbIn, _Out_writes_bytes_(*pcbOut) BYTE* pOut, _Inout_ DWORD* pcbOut, _Inout_ BOOL* pfDone, _In_opt_z_ LPCTSTR pszUserName, _In_opt_z_ LPCTSTR pszPassword, _In_opt_z_ LPCTSTR pszDomain);
  SECURITY_STATUS DoNTLMAuthentication(_In_opt_z_ LPCTSTR pszUsername, _In_opt_z_ LPCTSTR pszPassword, _In_opt_z_ LPCTSTR pszDomain);
  void ReleaseHandles();
  
//Member variables
  CredHandle m_hCred;                         //The credentials handle for the connection
  SecHandle  m_hContext;                      //The security handle for the connection 
  DWORD      m_dwBufferSize;                  //The size of the buffers we allocate for reading and writing via SSPI
};

#endif //#ifndef __PJNNTLMAUTH_H__
