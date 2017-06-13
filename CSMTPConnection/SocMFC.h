/*
Module : SocMFC.h
Purpose: Interface for a C++ wrapper class for sockets
Created: PJN / 05-08-1998

Copyright (c) 2002 - 2016 by PJ Naughter (Web: www.naughter.com, Email: pjna@naughter.com)

All rights reserved.

Copyright / Usage Details:

You are allowed to include the source code in any product (commercial, shareware, freeware or otherwise) 
when your product is released in binary form. You are allowed to modify the source code in any way you want 
except you cannot modify the copyright details at the top of each module. If you want to distribute source 
code with your application, then you are only allowed to distribute versions released by the author. This is 
to maintain a single distribution point for the source code. 

*/


/////////////////////////////// Macros / Defines //////////////////////////////

#pragma once

#ifndef __SOCMFC_H__
#define __SOCMFC_H__

#ifndef SOCKMFC_EXT_CLASS
#define SOCKMFC_EXT_CLASS
#endif //#ifndef SOCKMFC_EXT_CLASS

__if_not_exists(ADDRESS_FAMILY)
{
  typedef USHORT ADDRESS_FAMILY;
}

__if_not_exists(SOCKADDR_INET)
{
  typedef union _SOCKADDR_INET 
  {
      SOCKADDR_IN Ipv4;
      SOCKADDR_IN6 Ipv6;
      ADDRESS_FAMILY si_family;    
  } SOCKADDR_INET, *PSOCKADDR_INET;
}

#ifndef _In_
#define _In_
#endif //#ifndef _In_

#ifndef _Inout_
#define _Inout_
#endif //#ifndef _Inout_

#ifndef _Out_
#define _Out_
#endif //#ifndef _Out_

#ifndef _In_reads_bytes_opt_
#define _In_reads_bytes_opt_(X)
#endif //#ifndef _In_reads_bytes_opt_

#ifndef _Out_writes_bytes_to_
#define _Out_writes_bytes_to_(size,count)
#endif //#ifndef _Out_writes_bytes_to_

#ifndef _Out_writes_bytes_
#define _Out_writes_bytes_(size)
#endif //#ifndef _Out_writes_bytes_

#ifndef _Out_writes_bytes_opt_
#define _Out_writes_bytes_opt_(size)
#endif //#ifndef _Out_writes_bytes_opt_

#ifndef _Inout_opt_
#define _Inout_opt_
#endif //#ifndef _Inout_opt_

#ifndef _In_reads_bytes_
#define _In_reads_bytes_(size)
#endif //#ifndef _In_reads_bytes_

#ifndef _In_z_
#define _In_z_
#endif //#ifndef _In_z_

#ifndef __out_data_source
#define __out_data_source(src_sym)
#endif //#ifndef __out_data_source

#ifndef _Out_writes_bytes_to_opt_
#define _Out_writes_bytes_to_opt_(size,count)
#endif //#ifndef _Out_writes_bytes_to_opt_


////////////////////////////// Includes ///////////////////////////////////////

#ifndef _WINSOCK2API_
#pragma message("To avoid this message, please put winsock2.h in your pre compiled header (usually stdafx.h)")
#include <winsock2.h>
#endif //#ifndef _WINSOCK2API_

#ifndef CWSOCKET_MFC_EXTENSIONS

#ifndef _EXCEPTION_
#pragma message("To avoid this message, you should put exception in your pre compiled header (normally stdafx.h)")
#include <exception>
#endif //#ifndef _EXCEPTION_

#ifndef _STRING_
#pragma message("To avoid this message, please put string in your pre compiled header (normally stdafx.h)")
#include <string>
#endif //#ifndef _STRING_

#endif //#ifndef CWSOCKET_MFC_EXTENSIONS


////////////////////////////// Classes ////////////////////////////////////////

#ifdef CWSOCKET_MFC_EXTENSIONS
class SOCKMFC_EXT_CLASS CWSocketException : public CException
#else
class SOCKMFC_EXT_CLASS CWSocketException : public std::exception
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS
{
public:
//Constructors / Destructors
  CWSocketException(_In_ int nError);

//Methods
#ifdef CWSOCKET_MFC_EXTENSIONS
#ifdef _DEBUG
  virtual void Dump(CDumpContext& dc) const;
#endif //#ifdef _DEBUG
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

#if _MSC_VER >= 1700
  virtual BOOL GetErrorMessage(_Out_z_cap_(nMaxError) LPTSTR lpszError, _In_ UINT nMaxError,	_Out_opt_ PUINT pnHelpContext = NULL);
#else	
  virtual BOOL GetErrorMessage(__out_ecount_z(nMaxError) LPTSTR lpszError, __in UINT nMaxError, __out_opt PUINT pnHelpContext = NULL);
#endif //#if _MSC_VER >= 1700
  
#ifdef CWSOCKET_MFC_EXTENSIONS
  CString GetErrorMessage();
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

//Data members
  int m_nError;
};

class SOCKMFC_EXT_CLASS CWSocket
{
public:
//Constructors / Destructors
  CWSocket();
  virtual ~CWSocket();

//typedefs
#ifdef CWSOCKET_MFC_EXTENSIONS
  typedef CString String;
#else
  #ifdef _UNICODE
    typedef std::wstring String; 
  #else
    typedef std::string String;
  #endif //#ifdef _UNICODE
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

//Attributes
  void    Attach(_In_ SOCKET hSocket);
  SOCKET  Detach(); 
  void    GetPeerName(_Inout_ String& sPeerAddress, _Out_ UINT& nPeerPort, _In_ int nAddressToStringFlags = 0);
  void    GetPeerName(_Out_writes_bytes_to_(*pSockAddrLen, *pSockAddrLen) SOCKADDR* pSockaddr, _Inout_ int* pSockAddrLen);
  void    GetSockName(_Inout_ String& sSocketAddress, _Out_ UINT& nSocketPort, _In_ int nAddressToStringFlags = 0);
  void    GetSockName(_Out_writes_bytes_to_(*pSockAddrLen, *pSockAddrLen) SOCKADDR* pSockAddr, _Inout_ int* pSockAddrLen);
  void    SetSockOpt(_In_ int nOptionName, _In_reads_bytes_opt_(nOptionLen) const void* pOptionValue, _In_ int nOptionLen, _In_ int nLevel = SOL_SOCKET);
  void    GetSockOpt(_In_ int nOptionName, _Out_writes_bytes_(*pOptionLen) void* pOptionValue, _Inout_ int* pOptionLen, _In_ int nLevel = SOL_SOCKET);
  BOOL    IsCreated() const; 
  BOOL    IsReadible(_In_ DWORD dwTimeout);
  BOOL    IsWritable(_In_ DWORD dwTimeout);
  void    SetBindAddress(_In_ const String& sBindAddress) { m_sBindAddress = sBindAddress; };
  String  GetBindAddress() const { return m_sBindAddress; };

//Methods
  void    Create(_In_ BOOL bUDP = FALSE, _In_ BOOL bIPv6 = FALSE);
  void    Create(_In_ int nSocketType, _In_ int nProtocolType, _In_ int nAddressFormat);
  void    Accept(_Inout_ CWSocket& connectedSocket, _Out_writes_bytes_opt_(*pSockAddrLen) SOCKADDR* pSockAddr = NULL, _Inout_opt_ int* pSockAddrLen = NULL);
  void    CreateAndBind(_In_ UINT nSocketPort, _In_ int nSocketType = SOCK_STREAM, _In_ int nDefaultAddressFormat = AF_INET);
  void    Bind(_In_reads_bytes_(nSockAddrLen) const SOCKADDR* pSockAddr, _In_ int nSockAddrLen);
  void    Close();
  void    Connect(_In_reads_bytes_(nSockAddrLen) const SOCKADDR* pSockAddr, _In_ int nSockAddrLen);
  void    CreateAndConnect(_In_z_ LPCTSTR pszHostAddress, _In_ UINT nHostPort, _In_ int nSocketType = SOCK_STREAM, _In_ int nFamily = AF_UNSPEC, _In_ int nProtocolType = 0);
  void    CreateAndConnect(_In_z_ LPCTSTR pszHostAddress, _In_z_ LPCTSTR pszPortOrServiceName, _In_ int nSocketType = SOCK_STREAM, _In_ int nFamily = AF_UNSPEC, _In_ int nProtocolType = 0);
  void    Connect(_In_reads_bytes_(nSockAddrLen) const SOCKADDR* pSockAddr, _In_ int nSockAddrLen, _In_ DWORD dwTimeout, _In_ BOOL bResetToBlockingMode = TRUE);
  void    CreateAndConnect(_In_z_ LPCTSTR pszHostAddress, _In_ UINT nHostPort, _In_ DWORD dwTimeout, _In_ BOOL bResetToBlockingMode, _In_ int nSocketType = SOCK_STREAM);
  void    CreateAndConnect(_In_z_ LPCTSTR pszHostAddress, _In_z_ LPCTSTR pszPortOrServiceName, _In_ DWORD dwTimeout, _In_ BOOL bResetToBlockingMode, _In_ int nSocketType = SOCK_STREAM);
  void    IOCtl(_In_ long lCommand, _Inout_ DWORD* pArgument);
  void    Listen(_In_ int nConnectionBacklog = SOMAXCONN);
  int     Receive(_Out_writes_bytes_to_(nBufLen, return) __out_data_source(NETWORK) void* pBuf, _In_ int nBufLen, _In_ int nFlags = 0);
  int     ReceiveFrom(_Out_writes_bytes_to_(nBufLen, return) __out_data_source(NETWORK) void* pBuf, _In_ int nBufLen, _Out_writes_bytes_to_opt_(*pSockAddrLen, *pSockAddrLen) SOCKADDR* pSockAddr, _Inout_opt_ int* pSockAddrLen, _In_ int nFlags = 0);
  int     ReceiveFrom(_Out_writes_bytes_to_(nBufLen, return) __out_data_source(NETWORK) void* pBuf, _In_ int nBufLen, _Inout_ String& sSocketAddress, _Out_ UINT& nSocketPort, _In_ int nReceiveFromFlags = 0, _In_ int nAddressToStringFlags = 0);
  int     Send(_In_reads_bytes_(nBufLen) const void* pBuffer, _In_ int nBufLen, _In_ int nFlags = 0);
  int     SendTo(_In_reads_bytes_(nBufLen) const void* pBuf, _In_ int nBufLen, _In_reads_bytes_(nSockAddrLen) const SOCKADDR* pSockAddr, _In_ int nSockAddrLen, _In_ int nFlags = 0);
  int     SendTo(_In_reads_bytes_(nBufLen) const void* pBuf, _In_ int nBufLen, _In_ UINT nHostPort, _In_z_ LPCTSTR pszHostAddress = NULL, _In_ int nFlags = 0);
  void    ShutDown(_In_ int nHow = SD_SEND);

//Operators
  operator SOCKET();

//Static methods
  static void ThrowWSocketException(_In_ int nError = 0);
  static String AddressToString(_In_reads_bytes_(nSockAddrLen) const SOCKADDR* pSockAddr, _In_ int nSockAddrLen, _In_ int nFlags = 0, _Inout_opt_ UINT* pnSocketPort = NULL);
  static String AddressToString(_In_ const SOCKADDR_INET& sockAddr, _In_ int nFlags = 0, _Inout_opt_ UINT* pnSocketPort = NULL);

protected:
//Methods
  void _Connect(_In_z_ LPCTSTR pszHostAddress, _In_z_ LPCTSTR pszPortOrServiceName);
  void _Bind(_In_z_ LPCTSTR pszPortOrServiceName);

//Member variables
  SOCKET m_hSocket;
  String m_sBindAddress;
};

#endif //#ifndef __SOCMFC_H__
