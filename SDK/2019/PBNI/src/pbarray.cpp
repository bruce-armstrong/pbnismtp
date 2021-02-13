// ***********************************************************
// Copyright © 2018 Appeon Limited and its subsidiaries. All rights reserved.
//
//    Purpose  :        Definition for PBUnboundedArrayCreator
//***********************************************************************

#include "pbni.h"

//***************************************************************************
//	template <> PBUnboundedArrayCreator<pbstring>
//***************************************************************************
PBUnboundedArrayCreator<pbvalue_string>::PBUnboundedArrayCreator
(
	IPB_Session* session,
	pblong initialSize
)
:	d_session(session),
	d_array(0)
{
	d_array = session->NewUnboundedSimpleArray(pbvalue_string);
	if (initialSize > 0)
	{
	}
}

PBUnboundedArrayCreator<pbvalue_string>::~PBUnboundedArrayCreator()
{
}

void	PBUnboundedArrayCreator<pbvalue_string>::SetAt(pblong pos, pbstring v)
{
	pblong dim[1];
	dim[0] = pos;
	d_session->SetPBStringArrayItem(d_array, &dim[0], v);
}

void	PBUnboundedArrayCreator<pbvalue_string>::SetAt(pblong pos, LPCTSTR v)
{
	pblong dim[1];
	dim[0] = pos;
	d_session->SetStringArrayItem(d_array, &dim[0], v);
}

pbarray	PBUnboundedArrayCreator<pbvalue_string>::GetArray()
{
	return d_array;
}

//***************************************************************************
//	PBUnboundedObjectArrayCreator
//***************************************************************************
PBUnboundedObjectArrayCreator::PBUnboundedObjectArrayCreator
(
	IPB_Session* session,
	pbclass cls,
	pblong initialSize
)
:	d_session(session),
	d_array(0),
	d_class(cls)
{
	d_array = session->NewUnboundedObjectArray(cls);
	if (initialSize > 0)
	{
	}
}

PBUnboundedObjectArrayCreator::~PBUnboundedObjectArrayCreator()
{
}

void	PBUnboundedObjectArrayCreator::SetAt(pblong pos, pbobject obj)
{
	pblong dim[1];
	dim[0] = pos;
	d_session->SetObjectArrayItem(d_array, &dim[0], obj);
}

pbarray PBUnboundedObjectArrayCreator::GetArray()
{
	return d_array;
}

//***************************************************************************
//	template <> PBBoundedArrayCreator<pbstring>
//***************************************************************************
PBBoundedArrayCreator<pbvalue_string>::PBBoundedArrayCreator
(
	IPB_Session* session,
	pbuint dimensions,
	PBArrayInfo::ArrayBound* bounds
)
:	d_session(session),
	d_array(0)
{
	d_array = session->NewBoundedSimpleArray(pbvalue_string, dimensions, bounds);
}

PBBoundedArrayCreator<pbvalue_string>::~PBBoundedArrayCreator()
{
}

void	PBBoundedArrayCreator<pbvalue_string>::SetAt(pblong dim[], pbstring v)
{
	d_session->SetPBStringArrayItem(d_array, dim, v);
}

void	PBBoundedArrayCreator<pbvalue_string>::SetAt(pblong dim[], LPCTSTR v)
{
	d_session->SetStringArrayItem(d_array, dim, v);
}

pbarray	PBBoundedArrayCreator<pbvalue_string>::GetArray()
{
	return d_array;
}

//***************************************************************************
//	PBBoundedObjectArrayCreator
//***************************************************************************
PBBoundedObjectArrayCreator::PBBoundedObjectArrayCreator
(
	IPB_Session* session,
	pbclass cls,
	pbuint dimensions,
	PBArrayInfo::ArrayBound* bounds
)
:	d_session(session),
	d_class(cls),
	d_array(0)
{
	d_array = session->NewBoundedObjectArray(cls, dimensions, bounds);
}

PBBoundedObjectArrayCreator::~PBBoundedObjectArrayCreator()
{
}

void	PBBoundedObjectArrayCreator::SetAt(pblong dim[], pbobject obj)
{
	d_session->SetObjectArrayItem(d_array, dim, obj);
}

pbarray PBBoundedObjectArrayCreator::GetArray()
{
	return d_array;
}

//***************************************************************************
//	template<> PBArrayAccessor<pbvalue_string>
//***************************************************************************
PBArrayAccessor<pbvalue_string>::PBArrayAccessor
(
	IPB_Session* session,
	pbarray	array
)
:	d_session(session),
	d_array(array)
{
}

PBArrayAccessor<pbvalue_string>::~PBArrayAccessor()
{
}

void	PBArrayAccessor<pbvalue_string>::SetAt(pblong dim[], LPCTSTR str)
{
	pbstring pbstr = d_session->NewString(str);
	d_session->SetPBStringArrayItem(d_array, dim, pbstr);
}

void	PBArrayAccessor<pbvalue_string>::SetAt(pblong dim[], pbstring str)
{
	d_session->SetPBStringArrayItem(d_array, dim, str);
}

pbstring	PBArrayAccessor<pbvalue_string>::GetAt(pblong dim[], pbboolean& isNull)
{
	return d_session->GetStringArrayItem(d_array, dim, isNull);
}

//***************************************************************************
//	PBObjectArrayAccessor
//***************************************************************************
PBObjectArrayAccessor::PBObjectArrayAccessor
(
	IPB_Session* session,
	pbarray		array
)
:	d_session(session),
	d_array(array)
{
}

PBObjectArrayAccessor::~PBObjectArrayAccessor()
{
}

void	PBObjectArrayAccessor::SetAt(pblong dim[], pbobject obj)
{
	d_session->SetObjectArrayItem(d_array, dim, obj);
}

pbobject	PBObjectArrayAccessor::GetAt(pblong dim[], pbboolean& isNull)
{
	return d_session->GetObjectArrayItem(d_array, dim, isNull);
}
