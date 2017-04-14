/*
Module : CryptoWrappersCNGKey2.h
Purpose: Defines the interface for a C++ wrapper class for a Cryptography Next Generation
         NCRYPT_KEY_HANDLE
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

#ifndef __CRYPTOWRAPPERSCNGKEY2_H__
#define __CRYPTOWRAPPERSCNGKEY2_H__

#pragma comment(lib, "Ncrypt.lib")   //Automatically link in the CNG dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __NCRYPT_H__
#pragma message("To avoid this message, please put ncrypt.h in your pre compiled header (normally stdafx.h)")
#include <ncrypt.h>
#endif //#ifndef __NCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CNG NCRYPT_KEY_HANDLE
class CCNGKey2
{
public:
//Constructors / Destructors
  CCNGKey2() : m_hKey(NULL)
  {
  }
    
  CCNGKey2(_In_ CCNGKey2& key) : m_hKey(NULL)
  {
    Attach(key.Detach());
  }

  explicit CCNGKey2(_In_opt_ NCRYPT_KEY_HANDLE hKey) : m_hKey(hKey)
  {
  }

  ~CCNGKey2()
  {
    if (m_hKey != NULL)
      Free();
  }
  
//Methods
  CCNGKey2& operator=(_In_ CCNGKey2& key)
  {
    if (this != &key)
    {
      if (m_hKey != NULL)
        Free();
      Attach(key.Detach());
    }
    
    return *this;
  }

  SECURITY_STATUS Free()
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    SECURITY_STATUS result = NCryptFreeObject(m_hKey);
    m_hKey = NULL;
    return result;
  }

  SECURITY_STATUS Delete(_In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    SECURITY_STATUS status = NCryptDeleteKey(m_hKey, dwFlags);
    m_hKey = NULL; //Required because NCryptDeleteKey auto frees the m_hKey parameter
    return status;
  }

  NCRYPT_KEY_HANDLE Handle() const
  {
    return m_hKey;
  }

  void Attach(_In_opt_ NCRYPT_KEY_HANDLE hKey)
  {
    //Validate our parameters
    ATLASSERT(m_hKey == NULL);
    
    m_hKey = hKey;
  }

  NCRYPT_KEY_HANDLE Detach()
  {
    NCRYPT_KEY_HANDLE hKey = m_hKey;
    m_hKey = NULL;
    return hKey;
  }

  _Check_return_
  _Success_( return == 0 )
  SECURITY_STATUS GetProperty(_In_ LPCWSTR pszProperty, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PBYTE pbOutput, _In_ DWORD cbOutput, _Out_ DWORD* pcbResult, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return NCryptGetProperty(m_hKey, pszProperty, pbOutput, cbOutput, pcbResult, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS SetProperty(_In_ LPCWSTR pszProperty, _In_reads_bytes_(cbInput) PBYTE pbInput, _In_ DWORD cbInput, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return NCryptSetProperty(m_hKey, pszProperty, pbInput, cbInput, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS Export(_In_opt_ NCRYPT_KEY_HANDLE hExportKey, _In_ LPCWSTR pszBlobType, _In_opt_ NCryptBufferDesc* pParameterList,  _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PBYTE pbOutput, 
                         _In_ DWORD cbOutput, _Out_ DWORD* pcbResult, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return NCryptExportKey(m_hKey, hExportKey, pszBlobType, pParameterList, pbOutput, cbOutput, pcbResult, dwFlags);
  }

  _Success_(return != FALSE)
  BOOL ExportPublicKeyInfo(_In_ DWORD dwCertEncodingType, _Out_writes_bytes_to_opt_(*pcbInfo, *pcbInfo) PCERT_PUBLIC_KEY_INFO pInfo, _Inout_ DWORD* pcbInfo)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return CryptExportPublicKeyInfo(m_hKey, 0, dwCertEncodingType, pInfo, pcbInfo);
  }

  _Success_(return != FALSE)
  BOOL ExportPublicKeyInfoEx(_In_ DWORD dwCertEncodingType, _In_opt_ LPSTR pszPublicKeyObjId, _In_ DWORD dwFlags, _Out_writes_bytes_to_opt_(*pcbInfo, *pcbInfo) PCERT_PUBLIC_KEY_INFO pInfo, _Inout_ DWORD* pcbInfo)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return CryptExportPublicKeyInfoEx(m_hKey, 0, dwCertEncodingType, pszPublicKeyObjId, dwFlags, nullptr, pInfo, pcbInfo);
  }

  _Check_return_
  SECURITY_STATUS Decrypt(_In_reads_bytes_opt_(cbInput) PBYTE pbInput, _In_ DWORD cbInput, _In_opt_ VOID* pPaddingInfo, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PBYTE pbOutput,
                          _In_ DWORD cbOutput,_Out_ DWORD* pcbResult, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return NCryptDecrypt(m_hKey, pbInput, cbInput, pPaddingInfo, pbOutput, cbOutput, pcbResult, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS Encrypt(_In_reads_bytes_opt_(cbInput) PBYTE pbInput, _In_ DWORD cbInput, _In_opt_ VOID* pPaddingInfo,_Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PBYTE pbOutput,
                          _In_ DWORD cbOutput, _Out_ DWORD* pcbResult, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return NCryptEncrypt(m_hKey, pbInput, cbInput, pPaddingInfo, pbOutput, cbOutput, pcbResult, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS Finalize(_In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);
    
    return NCryptFinalizeKey(m_hKey, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS SignHash(_In_opt_ VOID* pPaddingInfo, _In_reads_bytes_(cbHashValue) PBYTE pbHashValue, _In_ DWORD cbHashValue, _Out_writes_bytes_to_opt_(cbSignature, *pcbResult) PBYTE pbSignature,
                           _In_ DWORD cbSignature, _Out_ DWORD* pcbResult, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return NCryptSignHash(m_hKey, pPaddingInfo, pbHashValue, cbHashValue, pbSignature, cbSignature, pcbResult, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS VerifySignature(_In_opt_ VOID* pPaddingInfo, _In_reads_bytes_(cbHashValue) PBYTE pbHashValue, _In_ DWORD cbHashValue,
                                  _In_reads_bytes_(cbSignature) PBYTE pbSignature, _In_ DWORD cbSignature, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return NCryptVerifySignature(m_hKey, pPaddingInfo, pbHashValue, cbHashValue, pbSignature, cbSignature, dwFlags);
  }

#if (NTDDI_VERSION >= NTDDI_WIN8)
  _Check_return_
  SECURITY_STATUS KeyDerivation(_In_opt_ NCryptBufferDesc* pParameterList, _Out_writes_bytes_to_(cbDerivedKey, *pcbResult) PUCHAR pbDerivedKey, _In_ DWORD cbDerivedKey, _Out_ DWORD* pcbResult, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return NCryptKeyDerivation(m_hKey, pParameterList, pbDerivedKey, cbDerivedKey, pcbResult, dwFlags);
  }
#endif //#if (NTDDI_VERSION >= NTDDI_WIN8)

  _Success_(return != FALSE)
  BOOL SignAndEncodeCertificate(_In_ DWORD dwCertEncodingType, _In_ LPCSTR lpszStructType, _In_ const void* pvStructInfo, _In_ PCRYPT_ALGORITHM_IDENTIFIER pSignatureAlgorithm, _Out_writes_bytes_to_opt_(*pcbEncoded, *pcbEncoded) BYTE* pbEncoded, _Inout_ DWORD* pcbEncoded)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return CryptSignAndEncodeCertificate(m_hKey, 0, dwCertEncodingType, lpszStructType, pvStructInfo, pSignatureAlgorithm, nullptr, pbEncoded, pcbEncoded);
  }

  _Success_(return != FALSE)
  BOOL SignCertificate(_In_ DWORD dwCertEncodingType, _In_reads_bytes_(cbEncodedToBeSigned) const BYTE* pbEncodedToBeSigned, _In_ DWORD cbEncodedToBeSigned, _In_ PCRYPT_ALGORITHM_IDENTIFIER pSignatureAlgorithm, _Out_writes_bytes_to_opt_(*pcbSignature, *pcbSignature) BYTE* pbSignature, _Inout_ DWORD* pcbSignature)
  {
    //Validate our parameters
    ATLASSERT(m_hKey != NULL);

    return CryptSignCertificate(m_hKey, 0, dwCertEncodingType, pbEncodedToBeSigned, cbEncodedToBeSigned, pSignatureAlgorithm, nullptr, pbSignature, pcbSignature);
  }

//Member variables
  NCRYPT_KEY_HANDLE m_hKey;

  friend class CCNGStorageProvider;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCNGKEY2_H__
