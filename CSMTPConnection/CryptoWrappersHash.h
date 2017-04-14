/*
Module : CryptoWrappersHash.h
Purpose: Defines the interface for a C++ wrapper class for a CryptoAPI HCRYPTHASH
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

#ifndef __CRYPTOWRAPPERSHASH_H__
#define __CRYPTOWRAPPERSHASH_H__

#pragma comment(lib, "Crypt32.lib")  //Automatically link in the Crypto dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, please put wincrypt.h in your pre compiled header (normally stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CryptoAPI HCRYPTHASH
class CHash
{
public:
//Constructors / Destructors
  CHash() : m_hHash(NULL)
  {
  }
  
  CHash(_In_ const CHash& hash) : m_hHash(NULL)
  {
    if (!hash.Duplicate(*this))
        m_hHash = NULL;
  }
  
  CHash(_In_ CHash&& hash) : m_hHash(NULL)
  {
    Attach(hash.Detach());
  }

  explicit CHash(_In_opt_ HCRYPTHASH hHash) : m_hHash(hHash)
  {
  }

  ~CHash()
  {
    if (m_hHash != NULL)
      Destroy();
  }
  
//Methods
  CHash& operator=(_In_ const CHash& hash)
  {
    if (this != &hash)
    {
      if (m_hHash != NULL)
        Destroy();
      if (!hash.Duplicate(*this))
        m_hHash = NULL;
    }
    
    return *this;
  }

  CHash& operator=(_In_ CHash&& hash)
  {
    if (m_hHash != NULL)
      Destroy();
    Attach(hash.Detach());
    
    return *this;
  }

  _Success_(return != FALSE)
  BOOL Destroy()
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    BOOL bResult = CryptDestroyHash(m_hHash);
    m_hHash = NULL;
    return bResult;
  }

  HCRYPTHASH Handle() const
  {
    return m_hHash;
  }

  void Attach(_In_opt_ HCRYPTHASH hHash)
  {
    //Validate our parameters
    ATLASSERT(m_hHash == NULL);
    
    m_hHash = hHash;
  }

  HCRYPTHASH Detach()
  {
    HCRYPTHASH hHash = m_hHash;
    m_hHash = NULL;
    return hHash;
  }

  _Success_(return != FALSE)
  BOOL Duplicate(_Inout_ CHash& hash) const
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);
    ATLASSERT(hash.m_hHash == NULL);

    return CryptDuplicateHash(m_hHash, nullptr, 0, &hash.m_hHash);
  }

  _Success_(return != FALSE)
  BOOL GetParam(_In_ DWORD dwParam, _Out_writes_bytes_to_opt_(*pdwDataLen, *pdwDataLen) BYTE* pbData, _Inout_ DWORD* pdwDataLen)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return CryptGetHashParam(m_hHash, dwParam, pbData, pdwDataLen, 0);
  }

  _Success_(return != FALSE)
  BOOL HashData(_In_reads_bytes_(dwDataLen) CONST BYTE* pbData, _In_ DWORD dwDataLen, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return CryptHashData(m_hHash, pbData, dwDataLen, dwFlags);
  }

  _Success_(return != FALSE)
  BOOL HashSessionKey(_In_ HCRYPTKEY hKey, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return CryptHashSessionKey(m_hHash, hKey, dwFlags);
  }

  _Success_(return != FALSE)
#if (NTDDI_VERSION >= NTDDI_WINXP)
  BOOL SetParam(_In_ DWORD dwParam, _In_ CONST BYTE* pbData)
#else
  BOOL SetParam(_In_ DWORD dwParam, _In_ BYTE* pbData)
#endif //#if (NTDDI_VERSION >= NTDDI_WINXP)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return CryptSetHashParam(m_hHash, dwParam, pbData, 0);
  }

  _Success_(return != FALSE)
  BOOL SignHash(_In_ DWORD dwKeySpec, _In_ DWORD dwFlags, _Out_writes_bytes_to_opt_(*pdwSigLen, *pdwSigLen) BYTE* pbSignature, _Inout_ DWORD* pdwSigLen)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return CryptSignHash(m_hHash, dwKeySpec, nullptr, dwFlags, pbSignature, pdwSigLen);
  }

  _Success_(return != FALSE)
  BOOL VerifySignature(_In_ BYTE* pbSignature, _In_ DWORD dwSigLen,_In_ HCRYPTKEY hPubKey, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return CryptVerifySignature(m_hHash, pbSignature, dwSigLen, hPubKey, nullptr, dwFlags);
  }

//Member variables
  HCRYPTHASH m_hHash;

  friend class CProvider;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSHASH_H__
