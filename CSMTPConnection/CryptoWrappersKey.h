/*
Module : CryptoWrappersKey.h
Purpose: Defines the interface for a C++ wrapper class for a CryptoAPI HCRYPTKEY
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

#ifndef __CRYPTOWRAPPERSKEY_H__
#define __CRYPTOWRAPPERSKEY_H__

#pragma comment(lib, "Crypt32.lib")  //Automatically link in the Crypto dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, please put wincrypt.h in your pre compiled header (normally stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CryptoAPI HCRYPTKEY
class CKey
{
public:
//Constructors / Destructors
  CKey() : m_hKey(NULL)
  {
  }
  
  CKey(_In_ CKey&& key) : m_hKey(NULL)
  {
    Attach(key.Detach());
  }
  
  CKey(_In_ const CKey& key) : m_hKey(NULL)
  {
    if (!key.Duplicate(*this))
      m_hKey = NULL;
  }

  explicit CKey(_In_opt_ HCRYPTKEY hKey) : m_hKey(hKey)
  {
  }

  ~CKey()
  {
    if (m_hKey != NULL)
      Destroy();
  }
  
//Methods
  CKey& operator=(_In_ const CKey& key)
  {
    if (this != &key)
    {
      if (m_hKey != NULL)
        Destroy();
      if (!key.Duplicate(*this))
        m_hKey = NULL;
    }
    
    return *this;
  }

  CKey& operator=(_In_ CKey&& key)
  {
    if (m_hKey != NULL)
      Destroy();
    Attach(key.Detach());
    
    return *this;
  }

  _Success_(return != FALSE)
  BOOL Destroy()
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    BOOL bResult = CryptDestroyKey(m_hKey);
    m_hKey = NULL;
    return bResult;
  }

  HCRYPTKEY Handle() const
  {
    return m_hKey;
  }

  void Attach(_In_opt_ HCRYPTKEY hKey)
  {
    //Validate our parameters
    ATLASSERT(m_hKey == NULL);
    
    m_hKey = hKey;
  }

  HCRYPTKEY Detach()
  {
    HCRYPTKEY hKey = m_hKey;
    m_hKey = NULL;
    return hKey;
  }

  _Success_(return != FALSE)
  BOOL Duplicate(_Inout_ CKey& key) const
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return CryptDuplicateKey(m_hKey, nullptr, 0, &key.m_hKey);
  }

  _Success_(return != FALSE)
  BOOL Export(_In_ HCRYPTKEY hExpKey, _In_ DWORD dwBlobType, _In_ DWORD dwFlags, _Out_writes_bytes_to_opt_(*pdwDataLen, *pdwDataLen) BYTE* pbData, _Inout_ DWORD* pdwDataLen)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return CryptExportKey(m_hKey, hExpKey, dwBlobType, dwFlags, pbData, pdwDataLen);
  }

  _Success_(return != FALSE)
  BOOL GetParam(_In_ DWORD dwParam, _Out_writes_bytes_to_opt_(*pdwDataLen, *pdwDataLen) BYTE* pbData, _Inout_ DWORD* pdwDataLen)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return CryptGetKeyParam(m_hKey, dwParam, pbData, pdwDataLen, 0);
  }

  _Success_(return != FALSE)
#if (NTDDI_VERSION >= NTDDI_WINXP)
  BOOL SetParam(_In_ DWORD dwParam, _In_ CONST BYTE* pbData, _In_ DWORD dwFlags)
#else
  BOOL SetParam(_In_ DWORD dwParam, _In_ BYTE* pbData, _In_ DWORD dwFlags)
#endif //#if (NTDDI_VERSION >= NTDDI_WINXP)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return CryptSetKeyParam(m_hKey, dwParam, pbData, dwFlags);
  }

  _Success_(0 != return)
  BOOL Decrypt(_In_ HCRYPTHASH hHash, _In_ BOOL Final, _In_ DWORD dwFlags, _Inout_updates_bytes_to_(*pdwDataLen, *pdwDataLen) BYTE* pbData, _Inout_ DWORD* pdwDataLen)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return CryptDecrypt(m_hKey, hHash, Final, dwFlags, pbData, pdwDataLen);
  }

  _Success_(0 != return) 
  BOOL Encrypt(_In_ HCRYPTHASH hHash, _In_ BOOL Final, _In_ DWORD dwFlags, _Inout_updates_bytes_to_opt_(dwBufLen, *pdwDataLen) BYTE* pbData, _Inout_ DWORD* pdwDataLen, _In_ DWORD dwBufLen)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return CryptEncrypt(m_hKey, hHash, Final, dwFlags, pbData, pdwDataLen, dwBufLen);
  }

//Member variables
  HCRYPTKEY m_hKey;

  friend class CProvider;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSKEY_H__
