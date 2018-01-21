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
//    Filename :        pbobject.cpp
//
//    Author   :        Steven Wang
//
//    Purpose  :        Implementation for PBObjectCreator
//
//***************************************************************************

#include "pbni.h"


PBObjectCreator::PBObjectCreator
(
	IPB_Session* session,
	LPCTSTR nvoName
)
: d_session(session)
{
	pbgroup group = d_session->FindGroup(nvoName, pbgroup_userobject);

	d_class = d_session->FindClass(group, nvoName);
}

PBObjectCreator::~PBObjectCreator()
{
}

pbobject PBObjectCreator::CreateObject()
{
	// assume d_session and d_class is valid here
	return d_session->NewObject(d_class);
}

