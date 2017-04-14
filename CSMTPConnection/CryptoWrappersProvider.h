/*
Module : CryptoWrappersProvider.h
Purpose: Defines the interface for a C++ wrapper class for a CryptoAPI HCRYPTPROV
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

#ifndef __CRYPTOWRAPPERSPROVIDER_H__
#define __CRYPTOWRAPPERSPROVIDER_H__

#pragma comment(lib, "Crypt32.lib")  //Automatically link in the Crypto dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, please put wincrypt.h in your pre compiled header (normally stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__
#include "CryptoWrappersDefaultContext.h"
#include "CryptoWrappersKey.h"
#include "CryptoWrappersHash.h"


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CryptoAPI HCRYPTPROV
class CProvider
{
public:
//Constructors / Destructors
  CProvider() : m_hProv(NULL)
  {
  }
  
  CProvider(_In_ const CProvider& prov) : m_hProv(NULL)
  {
    m_hProv = prov.Handle();
    if (m_hProv != NULL)
    {
        if (!AddRef())
          m_hProv = NULL;
    }
  }
  
  explicit CProvider(_In_opt_ HCRYPTPROV hProv) : m_hProv(hProv)
  {
  }

  ~CProvider()
  {
    if (m_hProv != NULL)
      Release();
  }
  
//Methods
  CProvider& operator=(_In_ const CProvider& prov)
  {
    if (this != &prov)
    {
      if (m_hProv != NULL)
        Release();
      m_hProv = prov.Handle();
      if (m_hProv != NULL)
      {
        if (!AddRef())
          m_hProv = NULL;
      }
    }
    
    return *this;
  }

  _Success_(return != FALSE)
  BOOL Acquire(_In_opt_ LPCTSTR pszContainer, _In_opt_ LPCTSTR pszProvider, _In_ DWORD dwProvType, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv == NULL);

    return CryptAcquireContext(&m_hProv, pszContainer, pszProvider, dwProvType, dwFlags);
  }

  _Success_(return != FALSE)
  BOOL Release()
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    BOOL bResult = CryptReleaseContext(m_hProv, 0);
    m_hProv = NULL;
    return bResult;
  }

  HCRYPTPROV Handle() const
  {
    return m_hProv;
  }

  void Attach(_In_opt_ HCRYPTPROV hProv)
  {
    //Validate our parameters
    ATLASSERT(m_hProv == NULL);
    
    m_hProv = hProv;
  }

  HCRYPTPROV Detach()
  {
    HCRYPTPROV hProv = m_hProv;
    m_hProv = NULL;
    return hProv;
  }

  _Success_(return != FALSE)
  BOOL AddRef()
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return CryptContextAddRef(m_hProv, nullptr, 0);
  }

  _Success_(return != FALSE)
  BOOL GetParam(_In_ DWORD dwParam, _Out_ BYTE* pbData, _Inout_ DWORD* pdwDataLen, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return CryptGetProvParam(m_hProv, dwParam, pbData, pdwDataLen, dwFlags);
  }

  _Success_(return != FALSE)
#if (NTDDI_VERSION >= NTDDI_WINXP)
  BOOL SetParam(_In_ DWORD dwParam, _In_ const BYTE* pbData, _In_ DWORD dwFlags)
#else
  BOOL SetParam(_In_ DWORD dwParam, _In_ BYTE* pbData, _In_ DWORD dwFlags)
#endif //#if (NTDDI_VERSION >= NTDDI_WINXP)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return CryptSetProvParam(m_hProv, dwParam, pbData, dwFlags);
  }

  _Success_(return != FALSE)
  BOOL InstallDefault(_In_ DWORD dwDefaultType, _In_opt_ const void* pvDefaultPara, _In_ DWORD dwFlags, _Inout_ CDefaultContext& ctx)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(ctx.Handle() == nullptr);

    HCRYPTDEFAULTCONTEXT hCtx = nullptr;
    BOOL bSuccess = CryptInstallDefaultContext(m_hProv, dwDefaultType, pvDefaultPara, dwFlags, nullptr, &hCtx);
    if (bSuccess)
        ctx.Attach(hCtx);

    return bSuccess;
  }

  _Success_(return != FALSE)
  BOOL DeriveKey(_In_ ALG_ID Algid, _In_ HCRYPTHASH hBaseData, _In_ DWORD dwFlags, _Inout_ CKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return CryptDeriveKey(m_hProv, Algid, hBaseData, dwFlags, &key.m_hKey);
  }

  _Success_(return != FALSE)
  BOOL GenKey(_In_ ALG_ID Algid, _In_ DWORD dwFlags, _Inout_ CKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return CryptGenKey(m_hProv, Algid, dwFlags, &key.m_hKey);
  }

  _Success_(return != FALSE)
  BOOL GenRandom(_In_ DWORD dwLen, _Inout_updates_bytes_(dwLen) BYTE* pbBuffer)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    
    return CryptGenRandom(m_hProv, dwLen, pbBuffer);
  }

  _Success_(return != FALSE)
  BOOL GetUserKey(_In_ DWORD dwKeySpec, _Inout_ CKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return CryptGetUserKey(m_hProv, dwKeySpec, &key.m_hKey);
  }

  _Success_(return != FALSE)
  BOOL ImportKey(_In_reads_bytes_(dwDataLen) CONST BYTE* pbData, _In_ DWORD dwDataLen, _In_ HCRYPTKEY hPubKey, _In_ DWORD dwFlags, _Inout_ CKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return CryptImportKey(m_hProv, pbData, dwDataLen, hPubKey, dwFlags, &key.m_hKey);
  }

  _Success_(return != FALSE)
  BOOL ImportPublicKeyInfo(_In_ DWORD dwCertEncodingType, _In_ PCERT_PUBLIC_KEY_INFO pInfo, _Inout_ CKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return CryptImportPublicKeyInfo(m_hProv, dwCertEncodingType, pInfo, &key.m_hKey);
  }

  _Success_(return != FALSE)
  BOOL ImportPublicKeyInfoEx(_In_ DWORD dwCertEncodingType, _In_ PCERT_PUBLIC_KEY_INFO pInfo, _In_ ALG_ID aiKeyAlg, _Inout_ CKey& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return CryptImportPublicKeyInfoEx(m_hProv, dwCertEncodingType, pInfo, aiKeyAlg, 0, nullptr, &key.m_hKey);
  }

  _Success_(return != FALSE)
  BOOL CreateHash(_In_ ALG_ID Algid, _In_ HCRYPTKEY hKey, _In_ DWORD dwFlags, _Inout_ CHash& hash)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(hash.m_hHash == NULL);

    return CryptCreateHash(m_hProv, Algid, hKey, dwFlags, &hash.m_hHash);
  }

  _Success_(return != FALSE)
  BOOL ExportPublicKeyInfo(_In_opt_ DWORD dwKeySpec, _In_ DWORD dwCertEncodingType, _Out_writes_bytes_to_opt_(*pcbInfo, *pcbInfo) PCERT_PUBLIC_KEY_INFO pInfo, _Inout_ DWORD* pcbInfo)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return CryptExportPublicKeyInfo(m_hProv, dwKeySpec, dwCertEncodingType, pInfo, pcbInfo);
  }

  _Success_(return != FALSE)
  BOOL ExportPublicKeyInfoEx(_In_opt_ DWORD dwKeySpec, _In_ DWORD dwCertEncodingType, _In_opt_ LPSTR pszPublicKeyObjId, _In_ DWORD dwFlags, _Out_writes_bytes_to_opt_(*pcbInfo, *pcbInfo) PCERT_PUBLIC_KEY_INFO pInfo, _Inout_ DWORD* pcbInfo)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return CryptExportPublicKeyInfoEx(m_hProv, dwKeySpec, dwCertEncodingType, pszPublicKeyObjId, dwFlags, nullptr, pInfo, pcbInfo);
  }

  _Success_(return != FALSE)
  BOOL SignAndEncodeCertificate(_In_opt_ DWORD dwKeySpec, _In_ DWORD dwCertEncodingType, _In_ LPCSTR lpszStructType, _In_ const void* pvStructInfo, _In_ PCRYPT_ALGORITHM_IDENTIFIER pSignatureAlgorithm, _Out_writes_bytes_to_opt_(*pcbEncoded, *pcbEncoded) BYTE *pbEncoded, _Inout_ DWORD *pcbEncoded)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return CryptSignAndEncodeCertificate(m_hProv, dwKeySpec, dwCertEncodingType, lpszStructType, pvStructInfo, pSignatureAlgorithm, nullptr, pbEncoded, pcbEncoded);
  }

  _Success_(return != FALSE)
  BOOL SignCertificate(_In_opt_ DWORD dwKeySpec, _In_ DWORD dwCertEncodingType, _In_reads_bytes_(cbEncodedToBeSigned) const BYTE* pbEncodedToBeSigned, _In_ DWORD cbEncodedToBeSigned, _In_ PCRYPT_ALGORITHM_IDENTIFIER pSignatureAlgorithm, _Out_writes_bytes_to_opt_(*pcbSignature, *pcbSignature) BYTE* pbSignature, _Inout_ DWORD* pcbSignature)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return CryptSignCertificate(m_hProv, dwKeySpec, dwCertEncodingType, pbEncodedToBeSigned, cbEncodedToBeSigned, pSignatureAlgorithm, nullptr, pbSignature, pcbSignature);
  }

//Member variables
  HCRYPTPROV m_hProv;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSPROVIDER_H__
