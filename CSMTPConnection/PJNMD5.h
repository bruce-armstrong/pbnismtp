/*
Module : PJNMD5.h
Purpose: Defines the interface for some MFC class which encapsulate calculating MD5 hashes and HMACS using the MS CryptoAPI
Created: PJN / 23-04-2005
History: PJN / 18-05-2005 1. Fixed a compiler warning when compiled using Visual Studio .NET 2003. Thanks to Alexey Kuznetsov
                          for reporting this issue.
         PJN / 29-09-2005 1. Format method now allows you to specify an uppercase or lower case string for the hash. This is 
                          necessary since the CRAM-MD5 authentication mechanism requires a lowercase MD5 hash. Thanks to 
                          Jian Peng for reporting this issue.
         PJN / 16-01-2015 1. Updated copright details
         PJN / 11-08-2016 1. Removed defunct comments about VC 6 in CPJNMD5 class.
                          2. Reworked the class to optionally compile without MFC. By default the class now uses STL 
                          classes and idioms but if you define CPJNMD5_MFC_EXTENSIONS the classes will revert back to the 
                          MFC behaviour.

Copyright (c) 2005 - 2017 by PJ Naughter (Web: www.naughter.com, Email: pjna@naughter.com)

All rights reserved.

Copyright / Usage Details:

You are allowed to include the source code in any product (commercial, shareware, freeware or otherwise) 
when your product is released in binary form. You are allowed to modify the source code in any way you want 
except you cannot modify the copyright details at the top of each module. If you want to distribute source 
code with your application, then you are only allowed to distribute versions released by the author. This is 
to maintain a single distribution point for the source code. 

*/


////////////////////////////// Macros / Defines ///////////////////////////////

#pragma once

#ifndef __PJNMD5_H__
#define __PJNMD5_H__


////////////////////////////// Includes ///////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, put wincrypt.h in your pre compiled header (usually stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__


////////////////////////////// Classes ////////////////////////////////////////

// A simple wrapper class which contains a MD5 hash i.e. a 16 byte blob
class CPJNMD5Hash
{
public:
//typedefs
#ifdef CPJNMD5_MFC_EXTENSIONS
  typedef CString String;
#else
#ifdef _UNICODE
  typedef std::wstring String;
#else
  typedef std::string String;
#endif //#ifdef _UNICODE
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS


//Constructors / Destructors
  CPJNMD5Hash()
  {
    memset(m_byHash, 0, sizeof(m_byHash));
  }

//methods
  operator BYTE*() 
  { 
    return m_byHash; 
  };

  String Format(_In_ BOOL bUppercase)
  {
    //What will be the return value from this method
  #ifdef CPJNMD5_MFC_EXTENSIONS
    CString sRet;
    LPTSTR pString = sRet.GetBuffer(33);
  #else
    String sRet;
    sRet.reserve(33);
    LPTSTR pString = &(sRet[0]);
  #endif //#ifdef CPJNMD5_MFC_EXTENSIONS

    for (int i=0; i<16; i++)
    {
      int nChar = (m_byHash[i] & 0xF0) >> 4;
      if (nChar <= 9)
        pString[i*2] = static_cast<TCHAR>(nChar + _T('0'));
      else
        pString[i*2] = static_cast<TCHAR>(nChar - 10 + (bUppercase ? _T('A') : _T('a')));

      nChar = m_byHash[i] & 0x0F;
      if (nChar <= 9)
        pString[i*2 + 1] = static_cast<TCHAR>(nChar + _T('0'));
      else
        pString[i*2 + 1] = static_cast<TCHAR>(nChar - 10 + (bUppercase ? _T('A') : _T('a')));
    }
    pString[32] = _T('\0');
#ifdef CPJNMD5_MFC_EXTENSIONS
    sRet.ReleaseBuffer();
  #endif //#ifdef CPJNMD5_MFC_EXTENSIONS

    return sRet;
  }

//Member variables
  BYTE m_byHash[16];
};


// Class which calculates MD5 Hashes and MD5 HMACS
class CPJNMD5
{
public:
//Constructors / Destructors
  CPJNMD5()
  {
    //Acquire the RSA CSP
    m_hProv = NULL;
    CryptAcquireContext(&m_hProv, NULL, NULL, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT | CRYPT_SILENT);
  }
  virtual ~CPJNMD5()
  {
    //Release the CSP
    if (m_hProv)
      CryptReleaseContext(m_hProv, 0);
  }


//Methods
  //Simple MD5 hash function which hashes a blob of data and returns
  //the results encapsulated in the "hash" parameter
  BOOL Hash(_In_reads_bytes_(dwDataSize) const BYTE* pbyData, _In_ DWORD dwDataSize, _Inout_ CPJNMD5Hash& hash)
  {
    //Call the Helper Hash function which returns a HCRYPTHASH handle
    //which contains the hash
    HCRYPTHASH hHash = NULL;
    BOOL bSuccess = Hash(pbyData, dwDataSize, hHash);
    if (bSuccess)
    {
      //Get the actual hash value
      DWORD dwHashSize = sizeof(hash);
      bSuccess = CryptGetHashParam(hHash, HP_HASHVAL, hash, &dwHashSize, 0);
    }
    //Free up the hash handle now that we are finished with it
    if (hHash)
      CryptDestroyHash(hHash);

    return bSuccess;
  }

  //Function which returns a hash encoded as a HCRYPTHASH. If hHash is
  //NULL on entry then a new HCRYPTHASH is created. If hHash is non-null
  //then this existing hash is used and the data to hash is included into
  //the current hash. The client is responsible for freeing the hash handle
  BOOL Hash(_In_reads_bytes_(dwDataSize) const BYTE* pbyData, DWORD dwDataSize, _Inout_ HCRYPTHASH& hHash)
  {
    //Create the hash object if required to
    if (hHash == NULL)
      if (!CryptCreateHash(m_hProv, CALG_MD5, 0, 0, &hHash)) 
        return FALSE;

    //Hash in the data
    return CryptHashData(hHash, pbyData, dwDataSize, 0);
  }

  //Function which calculates a keyed MD5 HMAC according to RFC 2104
  BOOL HMAC(_In_reads_bytes_(dwDataSize) const BYTE* pbyData, _In_ DWORD dwDataSize, _In_reads_bytes_(dwKeySize) const BYTE* pbyKey, _In_ DWORD dwKeySize, _Inout_ CPJNMD5Hash& hash)
  {
    //if key is longer than 64 bytes then reset it to the MD5 hash of the key
    const BYTE* pbyLocalKey = NULL;
    DWORD dwLocalKeySize = 0;
    CPJNMD5Hash keyHash;
    if (dwKeySize > 64)
    {
      if (!Hash(pbyKey, dwKeySize, keyHash))
        return FALSE;

      dwLocalKeySize = sizeof(keyHash);
      pbyLocalKey = keyHash;
    }
    else
    {
      dwLocalKeySize = dwKeySize;
      pbyLocalKey = pbyKey;  
    }

    //start out by storing key in pads
    BYTE k_ipad[64]; //KEY XORd with ipad
    memset(k_ipad, 0, sizeof(k_ipad));
    BYTE k_opad[64]; //KEY XORd with opad
    memset(k_opad, 0, sizeof(k_opad));
    memcpy_s(k_ipad, sizeof(k_ipad), pbyLocalKey, dwLocalKeySize);
    memcpy_s(k_opad, sizeof(k_opad), pbyLocalKey, dwLocalKeySize);

    //XOR key with ipad value
    for (int i=0; i<sizeof(k_ipad); i++) 
      k_ipad[i] ^= 0x36;

    //XOR key with opad
    for (int i=0; i<sizeof(k_opad); i++) 
      k_opad[i] ^= 0x5c;

    //perform inner MD5
    HCRYPTHASH hInnerHash = NULL;
  #pragma warning(suppress: 6385)
    if (!Hash(k_ipad, sizeof(k_ipad), hInnerHash))
    {
      //Free up the hash before we return
      if (hInnerHash)
        CryptDestroyHash(hInnerHash);

      return FALSE;
    }
    if (!Hash(pbyData, dwDataSize, hInnerHash))
    {
      //Free up the hash before we return
      CryptDestroyHash(hInnerHash);

      return FALSE;
    }

    //Get the inner hash result
    CPJNMD5Hash InnerHash;
    DWORD dwHashSize = sizeof(InnerHash);
    if (!CryptGetHashParam(hInnerHash, HP_HASHVAL, InnerHash, &dwHashSize, 0))
    {
      //Free up the hash before we return
      CryptDestroyHash(hInnerHash);

      return FALSE;
    }

    //Free up the inner hash now that we are finished with it
    CryptDestroyHash(hInnerHash);

    //Perform outter MD5
    HCRYPTHASH hOuterHash = NULL;
  #pragma warning(suppress: 6385)
    if (!Hash(k_opad, sizeof(k_opad), hOuterHash))
    {
      //Free up the hash before we return
      if (hOuterHash)
        CryptDestroyHash(hOuterHash);

      return FALSE;
    }
    if (!Hash(InnerHash, sizeof(InnerHash), hOuterHash))
    {
      //Free up the hash before we return
      CryptDestroyHash(hOuterHash);

      return FALSE;
    }

    //Finally get the hash result
    dwHashSize = sizeof(hash);
    BOOL bSuccess = CryptGetHashParam(hOuterHash, HP_HASHVAL, hash, &dwHashSize, 0);

    //Free up the hash before we return
    CryptDestroyHash(hOuterHash);

    return bSuccess;
  }

protected:
//Member variables
  HCRYPTPROV m_hProv;
};

#endif //#ifndef __PJNMD5_H__
