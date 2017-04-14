/*
Module : CryptoWrappersCNGHash.h
Purpose: Defines the interface for a C++ wrapper class for a Cryptography Next Generation
         BCRYPT_HASH_HANDLE
History: PJN / 01-08-2014 1. Initial public release
         PJN / 01-11-2015 1. Added a CCNGHash::ProcessMultiOperations method to encapsulate the new BCryptProcessMultiOperations
                          provided with Windows 8.1 Update.

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

#ifndef __CRYPTOWRAPPERSCNGHASH_H__
#define __CRYPTOWRAPPERSCNGHASH_H__

#pragma comment(lib, "Bcrypt.lib")   //Automatically link in the CNG dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __BCRYPT_H__
#pragma message("To avoid this message, please put bcrypt.h in your pre compiled header (normally stdafx.h)")
#include <bcrypt.h>
#endif //#ifndef __BCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CNG BCRYPT_HASH_HANDLE
class CCNGHash
{
public:
//Constructors / Destructors
  CCNGHash() : m_hHash(NULL)
  {
  }
  
  CCNGHash(_In_ const CCNGHash& hash) : m_hHash(nullptr)
  {
    if (!BCRYPT_SUCCESS(hash.Duplicate(*this)))
      m_hHash = nullptr;
  }
  
  CCNGHash(_In_ CCNGHash&& hash) : m_hHash(nullptr)
  {
    Attach(hash.Detach());
  }

  explicit CCNGHash(_In_opt_ BCRYPT_HASH_HANDLE hHash) : m_hHash(hHash)
  {
  }

  ~CCNGHash()
  {
    if (m_hHash != NULL)
      Destroy();
  }
  
//Methods
  CCNGHash& operator=(_In_ const CCNGHash& hash)
  {
    if (this != &hash)
    {
      if (m_hHash != nullptr)
        Destroy();
      if (!BCRYPT_SUCCESS(hash.Duplicate(*this)))
        m_hHash = nullptr;
    }
    
    return *this;
  }

  CCNGHash& operator=(_In_ CCNGHash&& hash)
  {
    if (m_hHash != NULL)
      Destroy();
    Attach(hash.Detach());
    
    return *this;
  }

  NTSTATUS Destroy()
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    NTSTATUS status = BCryptDestroyHash(m_hHash);
    m_hHash = NULL;
    return status;
  }

  BCRYPT_HASH_HANDLE Handle() const
  {
    return m_hHash;
  }

  void Attach(_In_opt_ BCRYPT_HASH_HANDLE hHash)
  {
    //Validate our parameters
    ATLASSERT(m_hHash == NULL);
    
    m_hHash = hHash;
  }

  BCRYPT_HASH_HANDLE Detach()
  {
    BCRYPT_HASH_HANDLE hHash = m_hHash;
    m_hHash = NULL;
    return hHash;
  }

  _Must_inspect_result_
  NTSTATUS Duplicate(_Inout_ CCNGHash& hash, _Out_writes_bytes_all_opt_(cbHashObject) PUCHAR pbHashObject = NULL, _In_ ULONG cbHashObject = 0) const
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);
    ATLASSERT(hash.m_hHash == NULL);

    return BCryptDuplicateHash(m_hHash, &hash.m_hHash, pbHashObject, cbHashObject, 0);
  }

  _Must_inspect_result_
  NTSTATUS HashData(_In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return BCryptHashData(m_hHash, pbInput, cbInput, 0);
  }

  _Must_inspect_result_
  NTSTATUS FinishHash(_Out_writes_bytes_all_(cbOutput) PUCHAR pbOutput,  _In_ ULONG cbOutput)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return BCryptFinishHash(m_hHash, pbOutput, cbOutput, 0);
  }

  _Must_inspect_result_
  NTSTATUS GetProperty(_In_ LPCWSTR pszProperty, _Out_writes_bytes_to_opt_(cbOutput, *pcbResult) PUCHAR pbOutput, _In_ ULONG cbOutput, _Out_ ULONG* pcbResult)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);
    
    return BCryptGetProperty(m_hHash, pszProperty, pbOutput, cbOutput, pcbResult, 0);
  }

  _Must_inspect_result_
  NTSTATUS SetProperty(_In_ LPCWSTR pszProperty, _In_reads_bytes_(cbInput) PUCHAR pbInput, _In_ ULONG cbInput, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);
    
    return BCryptSetProperty(m_hHash, pszProperty, pbInput, cbInput, dwFlags);
  }

#if (NTDDI_VERSION > NTDDI_WINBLUE || (NTDDI_VERSION == NTDDI_WINBLUE && defined(WINBLUE_KBSPRING14)))
  _Must_inspect_result_
  NTSTATUS ProcessMultiOperations(_In_ BCRYPT_MULTI_OPERATION_TYPE operationType, _In_reads_bytes_(cbOperations) PVOID pOperations, _In_ ULONG cbOperations, _In_ ULONG dwFlags)
  {
    //Validate our parameters
    ATLASSERT(m_hHash != NULL);

    return BCryptProcessMultiOperations(m_hHash, operationType, pOperations, cbOperations, dwFlags);
  }
#endif //#if (NTDDI_VERSION > NTDDI_WINBLUE || (NTDDI_VERSION == NTDDI_WINBLUE && defined(WINBLUE_KBSPRING14)))

//Member variables
  BCRYPT_HASH_HANDLE m_hHash;

  friend class CCNGProvider;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCNGHASH_H__
