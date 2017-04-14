/*
Module : CryptoWrappersOCSPResponseContext.h
Purpose: Defines the interface for a C++ wrapper class for a CryptoAPI PCCERT_SERVER_OCSP_RESPONSE_CONTEXT
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

#ifndef __CRYPTOWRAPPERSOCSPRESPONSECONTEXT_H__
#define __CRYPTOWRAPPERSOCSPRESPONSECONTEXT_H__

#pragma comment(lib, "Crypt32.lib")  //Automatically link in the Crypto dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, please put wincrypt.h in your pre compiled header (normally stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CryptoAPI PCCERT_SERVER_OCSP_RESPONSE_CONTEXT
#if (NTDDI_VERSION >= NTDDI_VISTA)
class COCSPResponseContext
{
public:
//Constructors / Destructors
  COCSPResponseContext() : m_pResponseContext(nullptr)
  {
  }
  
  COCSPResponseContext(_In_ const COCSPResponseContext& responseContext) : m_pResponseContext(nullptr)
  {
    m_pResponseContext = responseContext.Handle();
    AddRef();
  }

  COCSPResponseContext(_In_ COCSPResponseContext&& responseContext) : m_pResponseContext(nullptr)
  {
    Attach(responseContext.Detach());
  }
  
  explicit COCSPResponseContext(_In_opt_ PCCERT_SERVER_OCSP_RESPONSE_CONTEXT pResponseContext) : m_pResponseContext(pResponseContext)
  {
  }

  ~COCSPResponseContext()
  {
    if (m_pResponseContext != nullptr)
      Free();
  }
  
//Methods
  COCSPResponseContext& operator=(_In_ const COCSPResponseContext& responseContext)
  {
    if (this != &responseContext)
    {
      if (m_pResponseContext != nullptr)
        Free();
      m_pResponseContext = responseContext.Handle();
      AddRef();
    }
    
    return *this;
  }

  COCSPResponseContext& operator=(_In_ COCSPResponseContext&& responseContext)
  {
    if (this != &responseContext)
    {
        if (m_pResponseContext != nullptr)
          Free();
        Attach(responseContext.Detach());
    }
    
    return *this;
  }

  PCCERT_SERVER_OCSP_RESPONSE_CONTEXT Handle() const
  {
    return m_pResponseContext;
  }

  void Attach(_In_opt_ PCCERT_SERVER_OCSP_RESPONSE_CONTEXT pResponseContext)
  {
    //Validate our parameters
    ATLASSERT(m_pResponseContext == nullptr);
    
    m_pResponseContext = pResponseContext;
  }

  PCCERT_SERVER_OCSP_RESPONSE_CONTEXT Detach()
  {
    PCCERT_SERVER_OCSP_RESPONSE_CONTEXT pResponseContext = m_pResponseContext;
    m_pResponseContext = nullptr;
    return pResponseContext;
  }

  void Free()
  {
    //Validate our parameters
    ATLASSERT(m_pResponseContext != nullptr);

    CertFreeServerOcspResponseContext(m_pResponseContext);
    m_pResponseContext = NULL;
  }

  void AddRef()
  {
    //Validate our parameters
    ATLASSERT(m_pResponseContext != nullptr);

    CertAddRefServerOcspResponseContext(m_pResponseContext);
  }

//Member variables
  PCCERT_SERVER_OCSP_RESPONSE_CONTEXT m_pResponseContext;
};
#endif //#if (NTDDI_VERSION >= NTDDI_VISTA)


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSOCSPRESPONSECONTEXT_H__
