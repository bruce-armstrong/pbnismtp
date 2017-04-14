/*
Module : CryptoWrappersMessage.h
Purpose: Defines the interface for a C++ wrapper class for a CryptoAPI HCRYPTMSG
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

#ifndef __CRYPTOWRAPPERSMESSAGE_H__
#define __CRYPTOWRAPPERSMESSAGE_H__

#pragma comment(lib, "Crypt32.lib")  //Automatically link in the Crypto dll


////////////////////////// Includes ///////////////////////////////////////////

#ifndef __WINCRYPT_H__
#pragma message("To avoid this message, please put wincrypt.h in your pre compiled header (normally stdafx.h)")
#include <wincrypt.h>
#endif //#ifndef __WINCRYPT_H__


////////////////////////// Classes ////////////////////////////////////////////

namespace CryptoWrappers
{


//Encapsulates a CryptoAPI HCRYPTMSG
class CMessage
{
public:
//Constructors / Destructors
  CMessage() : m_hMessage(nullptr)
  {
  }

  CMessage(_In_ const CMessage& message) : m_hMessage(nullptr)
  {
    if (!message.Duplicate(*this))
      m_hMessage = nullptr;
  }
  
  CMessage(_In_ CMessage&& message) : m_hMessage(nullptr)
  {
    Attach(message.Detach());
  }

  explicit CMessage(_In_opt_ HCRYPTMSG hMessage) : m_hMessage(hMessage)
  {
  }

  ~CMessage()
  {
    if (m_hMessage != nullptr)
      Close();
  }
  
//Methods
  CMessage& operator=(_In_ const CMessage& message)
  {
    if (this != &message)
    {
      if (m_hMessage != nullptr)
        Close();
      if (!message.Duplicate(*this))
        m_hMessage = nullptr;
    }
    
    return *this;
  }

  CMessage& operator=(_In_ CMessage&& message)
  {
    if (m_hMessage != nullptr)
      Close();
    Attach(message.Detach());
    
    return *this;
  }
  
  _Success_(return != FALSE)
  BOOL Close()
  {
    //Validate our parameters
    ATLASSERT(m_hMessage != nullptr);

    BOOL bResult = CryptMsgClose(m_hMessage);
    m_hMessage = nullptr;
    return bResult;
  }

  HCRYPTMSG Handle() const
  {
    return m_hMessage;
  }

  void Attach(_In_opt_ HCRYPTMSG hMessage)
  {
    //Validate our parameters
    ATLASSERT(m_hMessage == nullptr);
    
    m_hMessage = hMessage;
  }

  HCRYPTMSG Detach()
  {
    HCRYPTMSG hMessage = m_hMessage;
    m_hMessage = nullptr;
    return hMessage;
  }

  _Success_(return != FALSE)
  BOOL Duplicate(_Inout_ CMessage& message) const
  {
    //Validate our parameters
    ATLASSERT(m_hMessage != nullptr);
    ATLASSERT(message.m_hMessage == nullptr);

    message.m_hMessage = CryptMsgDuplicate(m_hMessage);
    return (message.m_hMessage != nullptr);
  }

  _Must_inspect_result_
  BOOL OpenToDecode(_In_ DWORD dwMsgEncodingType, _In_ DWORD dwFlags, _In_ DWORD dwMsgType, _In_opt_ PCMSG_STREAM_INFO pStreamInfo)
  {
    //Validate our parameters
    ATLASSERT(m_hMessage == nullptr);

    m_hMessage = CryptMsgOpenToDecode(dwMsgEncodingType, dwFlags, dwMsgType, NULL, nullptr, pStreamInfo);
    return (m_hMessage != nullptr);
  }

  _Must_inspect_result_
  BOOL OpenToEncode(_In_ DWORD dwMsgEncodingType, _In_ DWORD dwFlags, _In_ DWORD dwMsgType, _In_ void const* pvMsgEncodeInfo, _In_opt_ LPSTR pszInnerContentObjID, _In_opt_ PCMSG_STREAM_INFO pStreamInfo)
  {
    //Validate our parameters
    ATLASSERT(m_hMessage == nullptr);

    m_hMessage = CryptMsgOpenToEncode(dwMsgEncodingType, dwFlags, dwMsgType, pvMsgEncodeInfo, pszInnerContentObjID, pStreamInfo);
    return (m_hMessage != nullptr);
  }

  _Success_(return != FALSE)
  BOOL Control(_In_ DWORD dwFlags, _In_ DWORD dwCtrlType, _In_opt_ void const* pvCtrlPara)
  {
    //Validate our parameters
    ATLASSERT(m_hMessage != nullptr);
    
    return CryptMsgControl(m_hMessage, dwFlags, dwCtrlType, pvCtrlPara);
  }

  _Success_(return != FALSE)
  BOOL GetParam(_In_ DWORD dwParamType, _In_ DWORD dwIndex, _Out_writes_bytes_to_opt_(*pcbData, *pcbData) void* pvData, _Inout_ DWORD* pcbData)
  {
    //Validate our parameters
    ATLASSERT(m_hMessage != nullptr);

    return CryptMsgGetParam(m_hMessage, dwParamType, dwIndex, pvData, pcbData);
  }

  _Success_(return != FALSE)
  BOOL Update(_In_reads_bytes_opt_(cbData) const BYTE* pbData, _In_ DWORD cbData, _In_ BOOL fFinal)
  {
    //Validate our parameters
    ATLASSERT(m_hMessage != nullptr);

    return CryptMsgUpdate(m_hMessage, pbData, cbData, fFinal);
  }

  _Success_(return != FALSE)
  BOOL Countersign(_In_ DWORD dwIndex, _In_ DWORD cCountersigners, _In_reads_(cCountersigners) PCMSG_SIGNER_ENCODE_INFO rgCountersigners)
  {
    //Validate our parameters
    ATLASSERT(m_hMessage != nullptr);

    return CryptMsgCountersign(m_hMessage, dwIndex, cCountersigners, rgCountersigners);
  }

//Member variables
  HCRYPTMSG m_hMessage;
};


}; //namespace CryptoWrappers

#endif //#ifndef __CRYPTOWRAPPERSMESSAGE_H__
