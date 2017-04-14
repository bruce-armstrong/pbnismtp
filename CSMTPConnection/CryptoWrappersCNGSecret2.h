/*
Module : CryptoWrappersCNGSecret2.h
Purpose: Defines the interface for a C++ wrapper class for a Cryptography Next Generation
         NCRYPT_SECRET_HANDLE
History: PJN / 01-08-2014 1. Initial public release

Copyright (c) 2014 - 2016 by PJ Naughter (Web: www.naughter.com, Email: pjna@naughter.com)

All rights reserved.

Copyright / Usage Details:

You are allowed to include the source code in any product (commercial, shareware, freeware or otherwise) 
when your product is released in binary form. You are allowed to modify the source code in any way you want 
except you cannot modify the copyright details at the top of each module. If you want to distribute source 
code with your application, then you are only allowed to distribute versions released by the author. This is 
to maintain a single distribution point for the source code. 

*/


////////////////////////// Macros / Defines ///////////////////////////////////

#pragma once

#ifndef __CRYPTOWRAPPERSCNGSECRET2_H__
#define __CRYPTOWRAPPERSCNGSECRET2_H__

#pragma comment(lib, "Ncrypt.lib")   //Automatically link in the CNG dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __NCRYPT_H__
#pragma message("To avoid this message, please put ncrypt.h in your pre compiled header (normally stdafx.h)")
#include <ncrypt.h>
#endif //#ifndef __NCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CNG NCRYPT_SECRET_HANDLE
class CCNGSecret2
{
public:
//Constructors / Destructors
  CCNGSecret2() : m_hSecret(NULL)
  {
  }
  
  CCNGSecret2(_In_ CCNGSecret2& secret) : m_hSecret(NULL)
  {
    Attach(secret.Detach());
  }
  
  explicit CCNGSecret2(_In_opt_ NCRYPT_SECRET_HANDLE hSecret) : m_hSecret(hSecret)
  {
  }

  ~CCNGSecret2()
  {
    if (m_hSecret != NULL)
      Free();
  }
  
//Methods
  _Must_inspect_result_
  SECURITY_STATUS Create(_In_ NCRYPT_KEY_HANDLE hPrivKey, _In_ NCRYPT_KEY_HANDLE hPubKey, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hSecret == NULL);

    return NCryptSecretAgreement(hPrivKey, hPubKey, &m_hSecret, dwFlags);
  }

  CCNGSecret2& operator=(_In_ CCNGSecret2& secret)
  {
    if (m_hSecret != NULL)
      Free();
    Attach(secret.Detach());
    
    return *this;
  }

  SECURITY_STATUS Free()
  {
    //Validate our parameters
    ATLASSERT(m_hSecret != NULL);

    SECURITY_STATUS status = NCryptFreeObject(m_hSecret);
    m_hSecret = NULL;
    return status;
  }

  NCRYPT_SECRET_HANDLE Handle() const
  {
    return m_hSecret;
  }

  void Attach(_In_opt_ NCRYPT_SECRET_HANDLE hSecret)
  {
    //Validate our parameters
    ATLASSERT(m_hSecret == NULL);
    
    m_hSecret = hSecret;
  }

  NCRYPT_SECRET_HANDLE Detach()
  {
    NCRYPT_SECRET_HANDLE hSecret = m_hSecret;
    m_hSecret = NULL;
    return hSecret;
  }

  _Check_return_
  NTSTATUS DeriveKey(_In_ LPCWSTR pwszKDF, _In_opt_ NCryptBufferDesc* pParameterList, _Out_writes_bytes_to_opt_(cbDerivedKey, *pcbResult) PBYTE pbDerivedKey,
                     _In_ DWORD cbDerivedKey,_Out_ DWORD* pcbResult, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hSecret != NULL);

    return NCryptDeriveKey(m_hSecret, pwszKDF, pParameterList, pbDerivedKey, cbDerivedKey, pcbResult, dwFlags);
  }

//Member variables
  NCRYPT_SECRET_HANDLE m_hSecret;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCNGSECRET2_H__
