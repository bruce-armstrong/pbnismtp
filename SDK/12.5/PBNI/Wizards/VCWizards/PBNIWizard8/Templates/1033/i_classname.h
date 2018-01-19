// I[!output SAFE_CLASS_NAME].h : header file for PBNI class interface
#ifndef I[!output UPPER_CLASS_NAME]_H
#define I[!output UPPER_CLASS_NAME]_H

#include <pbext.h>

// I[!output SAFE_CLASS_NAME] is a C++ interface for another PBNI Extension Module
// to interact with [!output SAFE_CLASS_NAME].
//
// This other PBNI Extension gets the I[!output SAFE_CLASS_NAME] pointer 
// by using the IPB_Session::GetNativeInterface(pbobject obj) method 
// (where "obj" refers to a [!output SAFE_CLASS_NAME] pbobject) and then casting 
// the return value to an I[!output SAFE_CLASS_NAME] pointer.
//
// e.g.
//
//   I[!output SAFE_CLASS_NAME]* pI[!output SAFE_CLASS_NAME] = NULL;
//   pI[!output SAFE_CLASS_NAME] = (I[!output SAFE_CLASS_NAME]*)(session->GetNativeInterface((pbobject)obj));
//
// After getting pI[!output SAFE_CLASS_NAME], we can call the methods of I[!output SAFE_CLASS_NAME],
//
// e.g.
// 
//   pI[!output SAFE_CLASS_NAME]->SampleInterfaceFunction();
//
[!if VISUAL_FLAG]
struct I[!output SAFE_CLASS_NAME] : public IPBX_VisualObject
[!else]
struct I[!output SAFE_CLASS_NAME] : public IPBX_NonVisualObject
[!endif]
{
  public :
    virtual long SampleInterfaceFunction() = 0;
};

#endif   // !defined I[!output UPPER_CLASS_NAME]_H
