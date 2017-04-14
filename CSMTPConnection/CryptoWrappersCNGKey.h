/*
Module : CryptoWrappersCNGKey.h
Purpose: Defines the interface for a C++ wrapper class for a Cryptography Next Generation
         BCRYPT_KEY_HANDLE
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

#ifndef __CRYPTOWRAPPERSCNGKEY_H__
#define __CRYPTOWRAPPERSCNGKEY_H__

#pragma comment(lib, "Bcrypt.lib")   //Automatically link in the CNG dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __BCRYPT_H__
#pragma message("To avoid this message, please put bcrypt.h in your pre compiled header (normally stdafx.h)")
#include <bcrypt.h>
#endif //#ifndef __BCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CNG BCRYPT_KEY_HANDLE
class CCNGKey
{
public:
//Constructors / Destructors
  CCNGKey() : m_hKey(NULL)
  {
  }
    
  CCNGKey(_In_ const CCNGKey& key) : m_hKey(nullptr)
  {
    if (!BCRYPT_SUCCESS(key.Duplicate(*this, nullptr, 0)))
      m_hKey = nullptr;
  }

  explicit CCNGKey(_In_opt_ BCRYPT_KEY_HANDLE hKey) : m_hKey(hKey)
  {
  }

  ~CCNGKey()
  {
    if (m_hKey != nullptr)
      Destroy();
  }
  
//Methods
  CCNGKey& operator=(_In_ const CCNGKey& key)
  {
    if (this != &key)
    {
      if (m_hKey != nullptr)
        Destroy();
      if (!BCRYPT_SUCCESS(key.Duplicate(*this, nullptr, 0)))
        m_hKey = nullptr;
    }
    
    return *this;
  }

  CCNGKey& operator=(_In_ CCNGKey&& key)
  {
    if (m_hKey != NULL)
      Destroy();
    Attach(key.Detach());
    
    return *this;
  }

  NTSTATUS Destroy()
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    NTSTATUS result = BCryptDestroyKey(m_hKey);
    m_hKey = NULL;
    return result;
  }

  BCRYPT_KEY_HANDLE Handle() const
  {
    return m_hKey;
  }

  void Attach(_In_opt_ BCRYPT_KEY_HANDLE hKey)
  {
    //Validate our parameters
    ATLASSERT(m_hKey == NULL);
    
    m_hKey = hKey;
  }

  BCRYPT_KEY_HANDLE Detach()
  {
    BCRYPT_KEY_HANDLE hKey = m_hKey;
    m_hKey = NULL;
    return hKey;
  }

  _Must_inspect_result_
  NTSTATUS Duplicate(_Inout_ CCNGKey& key, _Out_writes_bytes_all_opt_(cbKeyObject)  PUCHAR pbKeyObject, _In_ ULONG cbKeyObject) const
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return BCryptDuplicateKey(m_hKey, &key.m_hKey, pbKeyObject, cbKeyObject, 0);
  }

  _Must_inspect_result_
  NTSTATUS GetProperty(_In_ LPCWSTR pszProperty, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PUCHAR pbOutput, _In_ ULONG cbOutput, _Out_ ULONG* pcbResult)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return BCryptGetProperty(m_hKey, pszProperty, pbOutput, cbOutput, pcbResult, 0);
  }

  _Must_inspect_result_
  NTSTATUS SetProperty(_In_ LPCWSTR pszProperty, _In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return BCryptSetProperty(m_hKey, pszProperty, pbInput, cbInput, 0);
  }

  _Must_inspect_result_
  NTSTATUS Export(_In_opt_ BCRYPT_KEY_HANDLE hExportKey, _In_ LPCWSTR pszBlobType, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PUCHAR pbOutput, _In_ ULONG cbOutput, _Out_ ULONG* pcbResult)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return BCryptExportKey(m_hKey, hExportKey, pszBlobType, pbOutput, cbOutput, pcbResult, 0);
  }

#if (NTDDI_VERSION >= NTDDI_WIN7)
  _Success_(return != FALSE)
  BOOL ExportPublicKeyInfo(_In_ DWORD dwCertEncodingType, _In_opt_ LPSTR pszPublicKeyObjId, _In_ DWORD dwFlags, _Out_writes_bytes_to_opt_(*pcbInfo, *pcbInfo) PCERT_PUBLIC_KEY_INFO pInfo, _Inout_ DWORD* pcbInfo)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return CryptExportPublicKeyInfoFromBCryptKeyHandle(m_hKey, dwCertEncodingType, pszPublicKeyObjId, dwFlags, nullptr, pInfo, pcbInfo);
  }
#endif //#if (NTDDI_VERSION >= NTDDI_WIN7)

  _Must_inspect_result_
  NTSTATUS Decrypt(_In_reads_bytes_opt_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput, _In_opt_ VOID* pPaddingInfo, _Inout_updates_bytes_opt_(cbIV) PUCHAR pbIV, _In_ ULONG cbIV, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PUCHAR pbOutput,
               _In_ ULONG cbOutput, _Out_ ULONG* pcbResult, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return BCryptDecrypt(m_hKey, pbInput, cbInput, pPaddingInfo, pbIV, cbIV, pbOutput, cbOutput, pcbResult, dwFlags);
  }

  _Must_inspect_result_
  NTSTATUS Encrypt(_In_reads_bytes_opt_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput, _In_opt_ VOID* pPaddingInfo, _Inout_updates_bytes_opt_(cbIV) PUCHAR pbIV, _In_ ULONG cbIV,
                   _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PUCHAR pbOutput, _In_ ULONG cbOutput, _Out_ ULONG* pcbResult, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return BCryptEncrypt(m_hKey, pbInput, cbInput, pPaddingInfo, pbIV, cbIV, pbOutput, cbOutput, pcbResult, dwFlags);
  }

  _Must_inspect_result_
  NTSTATUS FinalizeKeyPair()
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return BCryptFinalizeKeyPair(m_hKey, 0);
  }

  _Must_inspect_result_
  NTSTATUS SignHash(_In_opt_ VOID* pPaddingInfo, _In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PUCHAR pbOutput, _In_ ULONG cbOutput, _Out_ ULONG* pcbResult, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return BCryptSignHash(m_hKey, pPaddingInfo, pbInput, cbInput, pbOutput, cbOutput, pcbResult, dwFlags);
  }

  _Must_inspect_result_
  NTSTATUS VerifySignature(_In_opt_ VOID* pPaddingInfo, _In_reads_bytes_(cbHash) PUCHAR pbHash, _In_ ULONG cbHash, _In_reads_bytes_(cbSignature) PUCHAR pbSignature, _In_ ULONG cbSignature, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return BCryptVerifySignature(m_hKey, pPaddingInfo, pbHash, cbHash, pbSignature, cbSignature, dwFlags);
  }

#if (NTDDI_VERSION >= NTDDI_WIN8)
  _Must_inspect_result_
  NTSTATUS KeyDerivation(_In_opt_ BCryptBufferDesc* pParameterList, _Out_ PUCHAR pbDerivedKey, _In_ ULONG cbDerivedKey, _Out_ ULONG* pcbResult, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return BCryptKeyDerivation(m_hKey, pParameterList, pbDerivedKey, cbDerivedKey, pcbResult, dwFlags);
  }
#endif //#if (NTDDI_VERSION >= NTDDI_WIN8)

//Member variables
  BCRYPT_KEY_HANDLE m_hKey;

  friend class CCNGProvider;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCNGKEY_H__
