/*
Module : CryptoWrappersCNGStorageProvider.h
Purpose: Defines the interface for a C++ wrapper class for a Cryptography Next Generation
         NCRYPT_PROV_HANDLE
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

#ifndef __CRYPTOWRAPPERSCNGSTORAGEPROVIDER_H__
#define __CRYPTOWRAPPERSCNGSTORAGEPROVIDER_H__

#pragma comment(lib, "Ncrypt.lib")   //Automatically link in the CNG dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __NCRYPT_H__
#pragma message("To avoid this message, please put ncrypt.h in your pre compiled header (normally stdafx.h)")
#include <ncrypt.h>
#endif //#ifndef __NCRYPT_H__
#include "CryptoWrappersCNGKey2.h"


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CNG NCRYPT_PROV_HANDLE
class CCNGStorageProvider
{
public:
//Constructors / Destructors
  CCNGStorageProvider() : m_hProv(NULL)
  {
  }
  
  CCNGStorageProvider(_In_ CCNGStorageProvider& prov) : m_hProv(NULL)
  {
    Attach(prov.Detach());
  }
  
  explicit CCNGStorageProvider(_In_opt_ NCRYPT_PROV_HANDLE hProv) : m_hProv(hProv)
  {
  }

  ~CCNGStorageProvider()
  {
    if (m_hProv != NULL)
      Free();
  }
  
//Methods
  CCNGStorageProvider& operator=(_In_ CCNGStorageProvider& prov)
  {
    if (this != &prov)
    {
      if (m_hProv != NULL)
        Free();
      Attach(prov.Detach());
    }
    
    return *this;
  }

  _Check_return_
  SECURITY_STATUS Open(_In_opt_ LPCWSTR pszProviderName)
  {
    //Validate our parameters
    ATLASSERT(m_hProv == NULL);

    return NCryptOpenStorageProvider(&m_hProv, pszProviderName, 0);
  }

  SECURITY_STATUS Free()
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    SECURITY_STATUS status = NCryptFreeObject(m_hProv);
    m_hProv = NULL;
    return status;
  }

  NCRYPT_PROV_HANDLE Handle() const
  {
    return m_hProv;
  }

  void Attach(_In_opt_ NCRYPT_PROV_HANDLE hProv)
  {
    //Validate our parameters
    ATLASSERT(m_hProv == NULL);
    
    m_hProv = hProv;
  }

  NCRYPT_PROV_HANDLE Detach()
  {
    NCRYPT_PROV_HANDLE hProv = m_hProv;
    m_hProv = NULL;
    return hProv;
  }

  _Check_return_
  SECURITY_STATUS EnumAlgorithms(_In_ DWORD dwAlgOperations, _Out_ DWORD* pdwAlgCount, _Out_ NCryptAlgorithmName** ppAlgList, _In_  DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return NCryptEnumAlgorithms(m_hProv, dwAlgOperations, pdwAlgCount, ppAlgList, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS EnumKeys(_Outptr_ NCryptKeyName** ppKeyName, _Inout_ PVOID* ppEnumState, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return NCryptEnumKeys(m_hProv, nullptr, ppKeyName, ppEnumState,dwFlags);
  }

  _Check_return_
  SECURITY_STATUS IsAlgSupported(_In_ LPCWSTR pszAlgId, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return NCryptIsAlgSupported(m_hProv, pszAlgId, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS NotifyChangeKey(_Inout_ HANDLE* phEvent, _In_ DWORD dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);

    return NCryptNotifyChangeKey(m_hProv, phEvent, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS ImportKey(_In_opt_ NCRYPT_KEY_HANDLE hImportKey, _In_ LPCWSTR pszBlobType, _In_opt_ NCryptBufferDesc* pParameterList,
                            _In_reads_bytes_(cbData) PBYTE pbData, _In_ DWORD cbData, _In_ DWORD dwFlags, _Inout_ CCNGKey2& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return NCryptImportKey(m_hProv, hImportKey, pszBlobType, pParameterList, &key.m_hKey, pbData, cbData, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS OpenKey(_In_ LPCWSTR pszKeyName, _In_opt_ DWORD dwLegacyKeySpec, _In_ DWORD dwFlags, _Inout_ CCNGKey2& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return NCryptOpenKey(m_hProv, &key.m_hKey, pszKeyName, dwLegacyKeySpec, dwFlags);
  }

  _Check_return_
  SECURITY_STATUS CreatePersistedKey(_In_ LPCWSTR pszAlgId, _In_opt_ LPCWSTR pszKeyName, _In_ DWORD dwLegacyKeySpec, _In_ DWORD dwFlags, _Inout_ CCNGKey2& key)
  {
    //Validate our parameters
    ATLASSERT(m_hProv != NULL);
    ATLASSERT(key.m_hKey == NULL);

    return NCryptCreatePersistedKey(m_hProv, &key.m_hKey, pszAlgId, pszKeyName, dwLegacyKeySpec, dwFlags);
  }

//Member variables
  NCRYPT_PROV_HANDLE m_hProv;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCNGSTORAGEPROVIDER_H__
