//**************************************************************************
//
//                            Copyright 2002
//                              Sybase Inc.
//
//								Copyright ?2002
//						Sybase, Inc. and its subsidiaries.
//	-	-	-	-	-	-	-	-	-	-	-	-	-	-	-	-	-	-	-
//	Sybase, Inc.("Sybase") claims copyright in this program and documentation
//	as an unpublished work, versions of which were first licensed on the date
//	indicated in the foregoing notice. This claim of copyright does not imply
//	waiver of Sybase's other rights.
// ------------------------------------------------------------------------
//
//    Filename :        pbfuninv.cpp
//
//    Author   :        Steven
//
//    Purpose  :        Implementation of PBGlobalFunctionInvoker and PBObjectFunctionInvoker
//
//***************************************************************************
#include "pbni.h"


//****************************************************************************
//	PBObjectFunctionInvoker
//****************************************************************************
PBObjectFunctionInvoker::PBObjectFunctionInvoker
(
	IPB_Session* session,
	pbobject obj,
	LPCTSTR methodName,
	LPCTSTR signature
) :
d_session(session),
d_object(obj)
{
		pbclass cls = d_session->GetClass(d_object);
		d_mid = d_session->GetMethodID(cls, methodName, PBRT_FUNCTION, signature);

		if (d_mid != kUndefinedMethodID)
			d_session->InitCallInfo(cls, d_mid, &d_ci);
}

PBObjectFunctionInvoker::PBObjectFunctionInvoker
(
	IPB_Session* session,
	pbobject obj,
	pbmethodID mid
) :
d_session(session),
d_object(obj),
d_mid(mid)
{
	pbclass cls = d_session->GetClass(d_object);
}

PBObjectFunctionInvoker::~PBObjectFunctionInvoker()
{
	d_session->FreeCallInfo(&d_ci);
}

pblong PBObjectFunctionInvoker::GetNumArgs() const
{
	return d_ci.pArgs->GetCount();
}

IPB_Value* PBObjectFunctionInvoker::GetArg(pblong i)
{
	return d_ci.pArgs->GetAt( (pbint)i );
}

IPB_Value* PBObjectFunctionInvoker::GetReturnValue()
{
	return d_ci.returnValue;
}

PBXRESULT PBObjectFunctionInvoker::Invoke()
{
	return d_session->InvokeObjectFunction(d_object, d_mid, &d_ci);
}

//****************************************************************************
//	PBGlobalFunctionInvoker
//****************************************************************************
PBGlobalFunctionInvoker::PBGlobalFunctionInvoker
(
	IPB_Session* session,
	LPCTSTR methodName,
	LPCTSTR signature
) :
d_session(session),
d_class(NULL),
d_mid(kUndefinedMethodID)
{
	bool bFound = false;

	if (!bFound)
	{
		pbgroup group = d_session->FindGroup(methodName, pbgroup_function);
		if (group)
			d_class = d_session->FindClass(group, methodName);
		if (d_class)
			d_mid = d_session->GetMethodID(d_class, methodName, PBRT_FUNCTION, signature);

		if (d_mid != kUndefinedMethodID)
			bFound = true;
	}

	// Then search system functions
	if (!bFound)
	{
		//d_class = d_session->GetSystemFunctionsClass();
		pbgroup sysGroup = d_session->GetSystemGroup();
		d_class = d_session->FindClass(sysGroup, (LPCTSTR)_TEXT("SystemFunctions"));
		if (d_class)
			d_mid = d_session->GetMethodID(d_class, methodName, PBRT_FUNCTION, signature);

		if (d_mid != kUndefinedMethodID)
			bFound = true;
	}

	if (bFound)
		d_session->InitCallInfo(d_class, d_mid, &d_ci);
}


PBGlobalFunctionInvoker::~PBGlobalFunctionInvoker()
{
	d_session->FreeCallInfo(&d_ci);
}

pblong PBGlobalFunctionInvoker::GetNumArgs() const
{
	return d_ci.pArgs->GetCount();
}

IPB_Value* PBGlobalFunctionInvoker::GetArg(pblong i)
{
	return d_ci.pArgs->GetAt( (pbint)i );
}

IPB_Value* PBGlobalFunctionInvoker::GetReturnValue()
{
	return d_ci.returnValue;
}

PBXRESULT PBGlobalFunctionInvoker::Invoke()
{
	return d_session->InvokeClassFunction(d_class, d_mid, &d_ci);
}

//****************************************************************************
//	PBEventTrigger
//****************************************************************************
PBEventTrigger::PBEventTrigger
(
	IPB_Session* session,
	pbobject obj,
	LPCTSTR methodName,
	LPCTSTR signature
) :
d_session(session),
d_object(obj)
{
	pbclass cls = d_session->GetClass(d_object);
	d_mid = d_session->GetMethodID(cls, methodName, PBRT_EVENT, signature);

	if (d_mid != kUndefinedMethodID)
		d_session->InitCallInfo(cls, d_mid, &d_ci);
}

PBEventTrigger::PBEventTrigger
(
	IPB_Session* session,
	pbobject obj,
	pbmethodID mid
) :
d_session(session),
d_object(obj),
d_mid(mid)
{
	pbclass cls = d_session->GetClass(d_object);
}

PBEventTrigger::~PBEventTrigger()
{
	d_session->FreeCallInfo(&d_ci);
}

pblong PBEventTrigger::GetNumArgs() const
{
	return d_ci.pArgs->GetCount();
}

IPB_Value* PBEventTrigger::GetArg(pblong i)
{
	return d_ci.pArgs->GetAt( (pbint)i );
}

IPB_Value* PBEventTrigger::GetReturnValue()
{
	return d_ci.returnValue;
}

PBXRESULT PBEventTrigger::Trigger()
{
	return d_session->TriggerEvent(d_object, d_mid, &d_ci);
}

