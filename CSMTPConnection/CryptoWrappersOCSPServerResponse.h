/*
Module : CryptoWrappersOCSPServerResponse.h
Purpose: Defines the interface for a C++ wrapper class for a CryptoAPI HCERT_SERVER_OCSP_RESPONSE
History: PJN / 01-08-2014 1. Initial public release
         PJN / 01-11-2015 1. Updated COCSPServerResponse::Open to optionally take a PCERT_SERVER_OCSP_RESPONSE_OPEN_PARA
                          as exposed in the Windows 10 SDK.

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

#ifndef __CRYPTOWRAPPERSOCSPSERVERRESPONSE_H__
#define __CRYPTOWRAPPERSOCSPSERVERRESPONSE_H__

#pragma comment(lib, "Crypt32.lib")  //Automatically link in the Crypto dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, please put wincrypt.h in your pre compiled header (normally stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__
#include "CryptoWrappersOCSPResponseContext.h"


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


#if (NTDDI_VERSION >= NTDDI_VISTA)
//Encapsulates a CryptoAPI HCERT_SERVER_OCSP_RESPONSE
class COCSPServerResponse
{
public:
//Constructors / Destructors
  COCSPServerResponse() : m_pServerResponse(nullptr)
  {
  }
  
  COCSPServerResponse(_In_ const COCSPServerResponse& serverResponse) : m_pServerResponse(nullptr)
  {
    m_pServerResponse = serverResponse.Handle();
    AddRef();
  }

  COCSPServerResponse(_In_ COCSPServerResponse&& serverResponse) : m_pServerResponse(nullptr)
  {
    Attach(serverResponse.Detach());
  }
  
  explicit COCSPServerResponse(_In_opt_ HCERT_SERVER_OCSP_RESPONSE pServerResponse) : m_pServerResponse(pServerResponse)
  {
  }

  ~COCSPServerResponse()
  {
    if (m_pServerResponse != nullptr)
      Close();
  }
  
//Methods
  COCSPServerResponse& operator=(_In_ const COCSPServerResponse& serverResponse)
  {
    if (this != &serverResponse)
    {
      if (m_pServerResponse != nullptr)
        Close();
      m_pServerResponse = serverResponse.Handle();
      AddRef();
    }
    
    return *this;
  }

  COCSPServerResponse& operator=(_In_ COCSPServerResponse&& serverResponse)
  {
    if (this != &serverResponse)
    {
        if (m_pServerResponse != nullptr)
          Close();
        Attach(serverResponse.Detach());
    }
    
    return *this;
  }

  HCERT_SERVER_OCSP_RESPONSE Handle() const
  {
    return m_pServerResponse;
  }

  void Attach(_In_opt_ HCERT_SERVER_OCSP_RESPONSE pServerResponse)
  {
    //Validate our parameters
    ATLASSERT(m_pServerResponse == nullptr);
    
    m_pServerResponse = pServerResponse;
  }

  HCERT_SERVER_OCSP_RESPONSE Detach()
  {
    HCERT_SERVER_OCSP_RESPONSE pServerResponse = m_pServerResponse;
    m_pServerResponse = nullptr;
    return pServerResponse;
  }

  __if_exists(CERT_SERVER_OCSP_RESPONSE_OPEN_PARA)
  {
    _Success_(return != FALSE)
    BOOL Open(_In_ PCCERT_CHAIN_CONTEXT pChainContext, _In_opt_ PCERT_SERVER_OCSP_RESPONSE_OPEN_PARA pOpenPara = nullptr)
    {
      //Validate our parameters
      ATLASSERT(m_pServerResponse == nullptr);

      m_pServerResponse = CertOpenServerOcspResponse(pChainContext, 0, pOpenPara);
      return (m_pServerResponse != nullptr);
    }
  }
  __if_not_exists(CERT_SERVER_OCSP_RESPONSE_OPEN_PARA)
  {
    _Success_(return != FALSE)
    BOOL Open(_In_ PCCERT_CHAIN_CONTEXT pChainContext)
    {
      //Validate our parameters
      ATLASSERT(m_pServerResponse == nullptr);

      m_pServerResponse = CertOpenServerOcspResponse(pChainContext, 0, nullptr);
      return (m_pServerResponse != nullptr);
    }
  }

  void Close()
  {
    //Validate our parameters
    ATLASSERT(m_pServerResponse != nullptr);

    CertCloseServerOcspResponse(m_pServerResponse, 0);
    m_pServerResponse = NULL;
  }

  void AddRef()
  {
    //Validate our parameters
    ATLASSERT(m_pServerResponse != nullptr);

    CertAddRefServerOcspResponse(m_pServerResponse);
  }

  COCSPResponseContext GetResponseContext()
  {
    //Validate our parameters
    ATLASSERT(m_pServerResponse != nullptr);

    return COCSPResponseContext(CertGetServerOcspResponseContext(m_pServerResponse, 0, nullptr));
  }

//Member variables
  HCERT_SERVER_OCSP_RESPONSE m_pServerResponse;
};
#endif //#if (NTDDI_VERSION >= NTDDI_VISTA)


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSOCSPSERVERRESPONSE_H__
