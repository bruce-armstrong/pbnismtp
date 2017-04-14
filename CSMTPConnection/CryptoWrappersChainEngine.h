/*
Module : CryptoWrappersChainEngine.h
Purpose: Defines the interface for a C++ wrapper class for a CryptoAPI HCERTCHAINENGINE
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

#ifndef __CRYPTOWRAPPERSCHAINENGINE_H__
#define __CRYPTOWRAPPERSCHAINENGINE_H__

#pragma comment(lib, "Crypt32.lib")  //Automatically link in the Crypto dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, please put wincrypt.h in your pre compiled header (normally stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CryptoAPI HCERTCHAINENGINE
class CChainEngine
{
public:
//Constructors / Destructors
  CChainEngine() : m_pChainEngine(nullptr)
  {
  }
  
  CChainEngine(_In_ CChainEngine& chainEngine) : m_pChainEngine(nullptr)
  {
    Attach(chainEngine.Detach());
  }
  
  explicit CChainEngine(_In_opt_ HCERTCHAINENGINE pChainEngine) : m_pChainEngine(pChainEngine)
  {
  }

  ~CChainEngine()
  {
    if (m_pChainEngine != nullptr)
      Free();
  }
  
//Methods
  CChainEngine& operator=(_In_ CChainEngine& chainEngine)
  {
    if (this != &chainEngine)
    {
      if (m_pChainEngine != nullptr)
        Free();
      Attach(chainEngine.Detach());
    }
    
    return *this;
  }

  HCERTCHAINENGINE Handle() const
  {
    return m_pChainEngine;
  }

  void Attach(_In_opt_ HCERTCHAINENGINE pChainEngine)
  {
    //Validate our parameters
    ATLASSERT(m_pChainEngine == nullptr);
    
    m_pChainEngine = pChainEngine;
  }

  HCERTCHAINENGINE Detach()
  {
    HCERTCHAINENGINE pChainEngine = m_pChainEngine;
    m_pChainEngine = nullptr;
    return pChainEngine;
  }

  _Success_(return != FALSE)
  BOOL Create(_In_ PCERT_CHAIN_ENGINE_CONFIG pConfig)
  {
    //Validate our parameters
    ATLASSERT(m_pChainEngine == nullptr);
    
    return CertCreateCertificateChainEngine(pConfig, &m_pChainEngine);
  }

  void Free()
  {
    //Validate our parameters
    ATLASSERT(m_pChainEngine != nullptr);

    CertFreeCertificateChainEngine(m_pChainEngine);
    m_pChainEngine = nullptr;
  }

  _Success_(return != FALSE)
  BOOL GetChain(_In_ PCCERT_CONTEXT pCertContext, _In_opt_ LPFILETIME pTime, _In_ HCERTSTORE hAdditionalStore, _In_ PCERT_CHAIN_PARA pChainPara, _In_ DWORD dwFlags, _Inout_ CChain& chain)
  {
    //Validate our parameters
    ATLASSERT(m_pChainEngine != nullptr);
    ATLASSERT(chain.m_pChain == nullptr);

    return CertGetCertificateChain(m_pChainEngine, pCertContext, pTime, hAdditionalStore, pChainPara, dwFlags, nullptr, &chain.m_pChain);
  }

//Member variables
  HCERTCHAINENGINE m_pChainEngine;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSCHAINENGINE_H__
