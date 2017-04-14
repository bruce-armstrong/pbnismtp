/*
Module : CryptoWrappersCNGProvider.h
Purpose: Defines the interface for a C++ wrapper class for a Cryptography Next Generation
         BCRYPT_ALG_HANDLE
History: PJN / 01-08-2014 1. Initial public release
         PJN / 01-11-2015 1. Updated SAL annotations for CCNGProvider::GenRandom to be consistent with the
                          annotations in the Windows 10 SDK.
                          2. Added a CCNGProvider::CreateMultiHash method to encapsulate the new BCryptCreateMultiHash
                          provided with Windows 8.1 Update.
                          3. Added a CCNGProvider::Hash method to encapsulate the new BCryptHash provided with Windows 10.
         PJN / 12-05-2016 1. Fixed a compile problem in CryptoWrappersCNGProvider.h when compiled in VC 2013 related to
                          the preprocessor define NTDDI_WINTHRESHOLD

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

#ifndef __CRYPTOWRAPPERSCNGPROVIDER_H__
#define __CRYPTOWRAPPERSCNGPROVIDER_H__

#pragma comment(lib, "Bcrypt.lib")   //Automatically link in the CNG dll

#ifndef NTDDI_WINTHRESHOLD
#define NTDDI_WINTHRESHOLD 0x0A000000
#endif //#ifndef NTDDI_WINTHRESHOLD


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __BCRYPT_H__
#pragma message("To avoid this message, please put bcrypt.h in your pre compiled header (normally stdafx.h)")
#include <bcrypt.h>
#endif //#ifndef __BCRYPT_H__
#include "CryptoWrappersCNGKey.h"
#include "CryptoWrappersCNGHash.h"


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CNG BCRYPT_ALG_HANDLE
class CCNGProvider
{
public:
//Constructors / Destructors
  CCNGProvider() : m_hProv(NULL)
  {
  }
  
  CCNGProvider(_In_ CCNGProvider& prov) : m_hProv(NULL)
  {
    Attach(prov.Detach());
  }
  
  explicit CCNGProvider(_In_opt_ BCRYPT_ALG_HANDLE hProv) : m_hProv(hProv)
  {
  }

  ~CCNGProvider()
  {
    if (m_hProv != NULL)
      Close();
  }
  
//Methods
  CCNGProvider& operator=(_In_ CCNGProvider& prov)
  {
    if (this != &prov)
    {
      if (m_hProv != NULL)
        Close();
      Attach(prov.Detach());
    }
    
    return *this;
  }

  _Must_inspect_result_
  NTSTATUS Open(LPCWSTR pszAlgId, _In_opt_ LPCWSTR pszImplementation, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv == NULL);

    return BCryptOpenAlgorithmProvider(&m_hProv, pszAlgId, pszImplementation, dwFlags);
  }

  NTSTATUS Close()
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    NTSTATUS status = BCryptCloseAlgorithmProvider(m_hProv, 0);
    m_hProv = NULL;
    return status;
  }

  BCRYPT_ALG_HANDLE Handle() const
  {
    return m_hProv;
  }

  void Attach(_In_opt_ BCRYPT_ALG_HANDLE hProv)
  {
    //Validate our parameters
    ATLASSERT(m_hProv == NULL);
    
    m_hProv = hProv;
  }

  BCRYPT_ALG_HANDLE Detach()
  {
    BCRYPT_ALG_HANDLE hProv = m_hProv;
    m_hProv = NULL;
    return hProv;
  }

  _Must_inspect_result_
  NTSTATUS ImportKey(_In_opt_ BCRYPT_KEY_HANDLE hImportKey, _In_ LPCWSTR pszBlobType, _Out_writes_bytes_all_opt_(cbKeyObject) PUCHAR pbKeyObject, _In_ ULONG cbKeyObject,
                     _In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput, _Inout_ CCNGKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return BCryptImportKey(m_hProv, hImportKey, pszBlobType, &key.m_hKey, pbKeyObject, cbKeyObject, pbInput, cbInput, 0);
  }

  _Must_inspect_result_
  NTSTATUS ImportKeyPair(_In_ LPCWSTR pszBlobType, _In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput, _In_ ULONG dwFlags, _Inout_ CCNGKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return BCryptImportKeyPair(m_hProv, nullptr, pszBlobType, &key.m_hKey, pbInput, cbInput, dwFlags);
  }

  _Must_inspect_result_
  NTSTATUS CreateHash(_Out_writes_bytes_all_opt_(cbHashObject) PUCHAR pbHashObject, _In_ ULONG cbHashObject, _In_reads_bytes_opt_(cbSecret) PUCHAR pbSecret, _In_ ULONG cbSecret, _In_ ULONG dwFlags, _Inout_ CCNGHash& hash)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(hash.m_hHash == NULL);

    return BCryptCreateHash(m_hProv, &hash.m_hHash, pbHashObject, cbHashObject, pbSecret, cbSecret, dwFlags);
  }

#if (NTDDI_VERSION > NTDDI_WINBLUE || (NTDDI_VERSION == NTDDI_WINBLUE && defined(WINBLUE_KBSPRING14)))
  _Must_inspect_result_
  NTSTATUS CreateMultiHash(_In_ ULONG nHashes, _Out_writes_bytes_all_opt_(cbHashObject) PUCHAR pbHashObject, _In_ ULONG cbHashObject, _In_reads_bytes_opt_(cbSecret) PUCHAR pbSecret, _In_ ULONG cbSecret, _In_ ULONG dwFlags, _Inout_ CCNGHash& hash)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(hash.m_hHash == NULL);

    return BCryptCreateMultiHash(m_hProv, &hash.m_hHash, nHashes, pbHashObject, cbHashObject, pbSecret, cbSecret, dwFlags);
  }
#endif //#if (NTDDI_VERSION > NTDDI_WINBLUE || (NTDDI_VERSION == NTDDI_WINBLUE && defined(WINBLUE_KBSPRING14)))

  _Must_inspect_result_
  NTSTATUS CreateSymmetricKey(_Out_writes_bytes_all_opt_(cbKeyObject) PUCHAR pbKeyObject, _In_ ULONG cbKeyObject, _In_reads_bytes_(cbSecret) PUCHAR pbSecret, _In_ ULONG cbSecret, _Inout_ CCNGKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return BCryptGenerateSymmetricKey(m_hProv, &key.m_hKey, pbKeyObject, cbKeyObject, pbSecret, cbSecret, 0);
  }

  _Must_inspect_result_
  NTSTATUS GenerateKeyPair(_In_ ULONG dwLength, _Inout_ CCNGKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return BCryptGenerateKeyPair(m_hProv, &key.m_hKey, dwLength, 0);
  }

  _Must_inspect_result_
  NTSTATUS GetProperty(_In_ LPCWSTR pszProperty, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PUCHAR pbOutput, _In_ ULONG cbOutput, _Out_ ULONG* pcbResult)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    
    return BCryptGetProperty(m_hProv, pszProperty, pbOutput, cbOutput, pcbResult, 0);
  }

  _Must_inspect_result_
  NTSTATUS SetProperty(_In_ LPCWSTR pszProperty, _In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    
    return BCryptSetProperty(m_hProv, pszProperty, pbInput, cbInput, 0);
  }

  _Must_inspect_result_
  NTSTATUS GenRandom(_Out_writes_bytes_(cbBuffer) PUCHAR pbBuffer, _In_ ULONG cbBuffer, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return BCryptGenRandom(m_hProv, pbBuffer, cbBuffer, dwFlags);
  }

#if (NTDDI_VERSION >= NTDDI_WINTHRESHOLD)
  NTSTATUS Hash(_In_reads_bytes_opt_(cbSecret) PUCHAR pbSecret, _In_ ULONG cbSecret, _In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput, _Out_writes_bytes_all_(cbOutput) PUCHAR pbOutput, _In_ ULONG cbOutput)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return BCryptHash(m_hProv, pbSecret, cbSecret, pbInput, cbInput, pbOutput, cbOutput);
  }
#endif //#if (NTDDI_VERSION >= NTDDI_WINTHRESHOLD)

//Member variables
  BCRYPT_ALG_HANDLE m_hProv;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCNGPROVIDER_H__
