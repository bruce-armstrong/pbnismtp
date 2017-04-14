/*
Module : CryptoWrappersCNGSecret.h
Purpose: Defines the interface for a C++ wrapper class for a Cryptography Next Generation
         BCRYPT_SECRET_HANDLE
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

#ifndef __CRYPTOWRAPPERSCNGSECRET_H__
#define __CRYPTOWRAPPERSCNGSECRET_H__

#pragma comment(lib, "Bcrypt.lib")   //Automatically link in the CNG dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __BCRYPT_H__
#pragma message("To avoid this message, please put bcrypt.h in your pre compiled header (normally stdafx.h)")
#include <bcrypt.h>
#endif //#ifndef __BCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CNG BCRYPT_SECRET_HANDLE
class CCNGSecret
{
public:
//Constructors / Destructors
  CCNGSecret() : m_hSecret(NULL)
  {
  }
  
  CCNGSecret(_In_ CCNGSecret& secret) : m_hSecret(NULL)
  {
    Attach(secret.Detach());
  }
  
  explicit CCNGSecret(_In_opt_ BCRYPT_SECRET_HANDLE hSecret) : m_hSecret(hSecret)
  {
  }

  ~CCNGSecret()
  {
    if (m_hSecret != NULL)
      Destroy();
  }
  
//Methods
  _Must_inspect_result_
  NTSTATUS Create(_In_ BCRYPT_KEY_HANDLE hPrivKey, _In_ BCRYPT_KEY_HANDLE hPubKey)
  {
    //Validate our parameters
    ATLASSERT(m_hSecret == NULL);

    return BCryptSecretAgreement(hPrivKey, hPubKey, &m_hSecret, 0);
  }

  CCNGSecret& operator=(_In_ CCNGSecret& secret)
  {
    if (m_hSecret != NULL)
      Destroy();
    Attach(secret.Detach());
    
    return *this;
  }

  NTSTATUS Destroy()
  {
    //Validate our parameters
    ATLASSERT(m_hSecret != NULL);

    NTSTATUS status = BCryptDestroySecret(m_hSecret);
    m_hSecret = NULL;
    return status;
  }

  BCRYPT_SECRET_HANDLE Handle() const
  {
    return m_hSecret;
  }

  void Attach(_In_opt_ BCRYPT_SECRET_HANDLE hSecret)
  {
    //Validate our parameters
    ATLASSERT(m_hSecret == NULL);
    
    m_hSecret = hSecret;
  }

  BCRYPT_SECRET_HANDLE Detach()
  {
    BCRYPT_SECRET_HANDLE hSecret = m_hSecret;
    m_hSecret = nullptr;
    return hSecret;
  }

  _Must_inspect_result_
  NTSTATUS DeriveKey(_In_ LPCWSTR pwszKDF, _In_opt_ BCryptBufferDesc* pParameterList, _Out_writes_bytes_to_opt_(cbDerivedKey, *pcbResult) PUCHAR pbDerivedKey, _In_ ULONG cbDerivedKey, _Out_ ULONG* pcbResult, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hSecret != NULL);

    return BCryptDeriveKey(m_hSecret, pwszKDF, pParameterList, pbDerivedKey, cbDerivedKey, pcbResult, dwFlags);
  }

//Member variables
  BCRYPT_SECRET_HANDLE m_hSecret;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCNGSECRET_H__
