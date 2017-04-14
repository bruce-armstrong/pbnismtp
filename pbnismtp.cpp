#include "CSMTPConnection\stdafx.h"
#include <pbext.h>
#include "pbnismtp.h"
#include "CSMTPConnection\PJNSMTP.h"
#include <tchar.h>

BOOL APIENTRY DllMain(
	HANDLE hModule,
    DWORD  reasonForCall,
    LPVOID lpReserved
)
{
	switch( reasonForCall )
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

PBXEXPORT LPCTSTR PBXCALL PBX_GetDescription()
{
	static const TCHAR desc[] =
	{
		_T ( "class n_cpp_smtp from nonvisualobject\n" )
		_T ( "function int Send ( )\n" )
		_T ( "subroutine SetMessage ( string pbmessage )\n" )
		_T ( "subroutine SetMessage ( string pbmessage, boolean pbHTML )\n" )
		_T ( "subroutine SetRecipientEmail ( string pbrecipientemail )\n" )
		_T ( "subroutine SetCCRecipientEmail ( string pbCCrecipientemail )\n" )
		_T ( "subroutine SetBCCRecipientEmail ( string pbBCCrecipientemail )\n" )
		_T ( "subroutine SetSenderEmail ( string pbsenderemail )\n" )
		_T ( "subroutine SetSMTPServer ( string pbsmtpserver )\n" )
		_T ( "subroutine SetSubject ( string pbsubject )\n" )
		_T ( "subroutine SetAttachment ( string pbattachment )\n" )
		_T ( "subroutine SetErrorMessagesOn ( )\n" )
		_T ( "subroutine SetErrorMessagesOff ( )\n" )
		_T ( "subroutine SetCharSet ( string pbcharset )\n" )
		_T ( "subroutine SetUsernamePassword ( string pbusername, string pbpassword )\n" )
		_T ( "subroutine SetNTMLAuthentication ( )\n" )
		_T ( "subroutine SetPort ( integer port )\n" )
		_T ( "subroutine SetSSL ( )\n" )
		_T ( "subroutine SetTLS ( )\n" )
		_T ( "end class\n" )
	};
	return (LPCTSTR)desc ;
}

PBXEXPORT PBXRESULT PBXCALL PBX_CreateNonVisualObject(
	IPB_Session*			pbsession, 
	pbobject				pbobj, 
	LPCTSTR					className, 
	IPBX_NonVisualObject	**obj
)
{

	LPCTSTR myclassname = _T ( "n_cpp_smtp" ) ;

	if ( _tcscmp( className, myclassname ) == 0 )
	{
		*obj = new CSMTP() ;
		return PBX_OK ;
	} ;

	*obj = NULL ;
	return PBX_E_NO_SUCH_CLASS ;

}

enum MethodIDs
{
	SEND = 0,
	SETMESSAGE = 1,
	SETMESSAGEHTML = 2,
	SETRECIPIENTEMAIL = 3,
	SETCCRECIPIENTEMAIL = 4,
	SETBCCRECIPIENTEMAIL = 5,
	SETSENDEREMAIL = 6,
	SETSMTPSERVER = 7,
	SETSUBJECT = 8,
	SETATTACHMENT = 9,
	SETERRORMESSAGESON = 10,   
	SETERRORMESSAGESOFF = 11,
	SETMSGCHARSET = 12,
	SETUSERNAMEPASSWORD = 13,
	SETNTLMAUTHENTICATION = 14,
	SETPORT = 15,
	SETSSL = 16,
	SETTLS = 17,
	ENTRY_COUNT
};

DWORD CSMTP::m_dwRefCount = 0;

CSMTP::CSMTP()
{	
	WSADATA wd;

	if( ++m_dwRefCount == 1 )
	{
		::WSAStartup( 0x0101, &wd );
	}

	MessageArg = NULL ;
	SenderEmailArg = NULL ;
	RecipientEmailArg = NULL ;
	CCRecipientEmailArg = NULL ;
	BCCRecipientEmailArg = NULL ;
	SMTPServerArg = NULL ;
	SubjectArg = NULL ;
	AttachmentArg = NULL ;
	CharSetArg = NULL ;
	UsernameArg = NULL ;
	PasswordArg = NULL ;

	Message = NULL;  
	SenderEmail = NULL;
	RecipientEmail = NULL;
	CCRecipientEmail = NULL;
	BCCRecipientEmail = NULL;
	SMTPServer = NULL;
	Subject = NULL; 
	Attachment = NULL;
	Username = NULL ;
	Password = NULL ;
	CharSet = _T("iso-8859-1");

	ErrorMessageBoxesOn = FALSE;    /* initialize to off  */	

	HTMLBody = FALSE ;

	NTLMAuthentication = FALSE;
	Port = 25 ;
	SSL = FALSE ;
	TLS = FALSE ;

}

CSMTP::~CSMTP(void)
{
	if( --m_dwRefCount == 0 )
	{
		::WSACleanup();
	}
}

int CSMTP::Send()
{

	TCHAR	title[] = _T ( "Debug" ) ;
	CPJNSMTPConnection connection ;
	CPJNSMTPMessage msg ;
	CPJNSMTPConnection::ConnectionType ConnectType = CPJNSMTPConnection::PlainText ;

	try
	{

		if ( SenderEmail == NULL )
		{
			return INVALIDSENDEREMAIL ;
		}

		if ( RecipientEmail == NULL && CCRecipientEmail == NULL && BCCRecipientEmail == NULL)
		{
			return INVALIDRECIPIENTEMAIL ;
		}

		if ( SMTPServer == NULL )
		{
			return INVALIDSMTPSERVER ;
		}

		if ( Subject == NULL )
		{
			Subject = _T ( "" ) ;
		}
			
		if ( Message == NULL )
		{
			Message = _T ( "" ) ;
		}

		if ( TLS )
		{
			ConnectType = CPJNSMTPConnection::STARTTLS ;
		}
		else
		{
			if ( SSL )
			{
				ConnectType = CPJNSMTPConnection::SSL_TLS ;
			}
		}

		//Connect to SMTP server
		#ifdef _DEBUG
			MessageBox( NULL, _T ( "Connecting to SMTP Server" ), title, MB_ICONEXCLAMATION | MB_OK );
		#endif
		if ( NTLMAuthentication )
		{
			connection.Connect( SMTPServer, CPJNSMTPConnection::AUTH_NTLM, Username, Password, Port, ConnectType ) ;
		}
		else if ( Username == NULL )
		{
			connection.Connect ( SMTPServer, CPJNSMTPConnection::AUTH_NONE, Username, Password, Port, ConnectType ) ;
		}
		else
		{
			connection.Connect ( SMTPServer, CPJNSMTPConnection::AUTH_AUTO, Username, Password, Port, ConnectType ) ;
		}
	}
	catch(CPJNSMTPException* pEx)
	{
		if (ErrorMessageBoxesOn == TRUE)
		{
			//Display the error
			CString sMsg;
			sMsg.Format(_T("Could not connect to SMTP server: Error:%08X\nDescription:%s"), pEx->m_hr, pEx->m_sLastResponse);
			MessageBox ( NULL, sMsg, title, MB_ICONEXCLAMATION | MB_OK);
		}
		pEx->Delete();
		return SMTPCONNECTFAILED ;
	}

	try
	{

		//Set the Charset
		msg.SetCharset ( (CString)CharSet ) ;

		//Say who the email is from
		CPJNSMTPAddress address ( (CString)SenderEmail ) ;
		msg.m_From = address ;

		//Say who the email is to
		msg.ParseMultipleRecipients ( (CString)RecipientEmail, msg.m_To ) ;
	}
	catch(CPJNSMTPException* pEx)
	{
		if (ErrorMessageBoxesOn == TRUE)
		{
			//Display the error
			CString sMsg;
			sMsg.Format(_T("Could not add TO recipient: Error:%08X\nDescription:%s"), pEx->m_hr, pEx->m_sLastResponse);
			MessageBox ( NULL, sMsg, title, MB_ICONEXCLAMATION | MB_OK);
		}
		pEx->Delete();
		return INVALIDRECIPIENTEMAIL ;
	}
	
	try
	{
		//Say who it's going to for cc
		if ( CCRecipientEmail != NULL )
		{
			msg.ParseMultipleRecipients ( (CString)CCRecipientEmail, msg.m_CC ) ;
		}
	}
	catch(CPJNSMTPException* pEx)
	{
		if (ErrorMessageBoxesOn == TRUE)
		{
			//Display the error
			CString sMsg;
			sMsg.Format(_T("Could not add CC recipient: Error:%08X\nDescription:%s"), pEx->m_hr, pEx->m_sLastResponse);
			MessageBox ( NULL, sMsg, title, MB_ICONEXCLAMATION | MB_OK);
		}
		pEx->Delete();
		return INVALIDRECIPIENTEMAIL ;
	}

	try
	{
		//Say who it's going to for bcc
		if ( BCCRecipientEmail != NULL )
		{
			msg.ParseMultipleRecipients ( (CString)BCCRecipientEmail, msg.m_BCC ) ;
		}
	}
	catch(CPJNSMTPException* pEx)
	{
		if (ErrorMessageBoxesOn == TRUE)
		{
			//Display the error
			CString sMsg;
			sMsg.Format(_T("Could not add BCC recipient: Error:%08X\nDescription:%s"), pEx->m_hr, pEx->m_sLastResponse);
			MessageBox ( NULL, sMsg, title, MB_ICONEXCLAMATION | MB_OK);
		}
		pEx->Delete();
		return INVALIDRECIPIENTEMAIL ;
	}

	try
	{
		//Set the subject and message
		msg.m_sSubject = (CString)Subject ; 
		msg.m_sXMailer = _T ( "PBNISMTP 3.0" ) ;
		if ( HTMLBody )
		{
			msg.SetMime(TRUE) ;
			msg.AddHTMLBody ( (CString)Message, _T ( "" ) ) ;
		}
		else
		{
			msg.AddTextBody ( (CString)Message ) ;
		}

		//Add any attachment
		if ( Attachment != NULL )
		{
			msg.AddMultipleAttachments ( (CString)Attachment ) ;
		}

		//Now send the message
		connection.SendMessage ( msg ) ;
	}
	catch(CPJNSMTPException* pEx)
	{
		if (ErrorMessageBoxesOn == TRUE)
		{
			//Display the error
			CString sMsg;
			sMsg.Format(_T("Could not send message: Error:%08X\nDescription:%s"), pEx->m_hr, pEx->m_sLastResponse);
			MessageBox ( NULL, sMsg, title, MB_ICONEXCLAMATION | MB_OK);
		}
		pEx->Delete();
		return MESSAGESENDERROR ;
	}

	return 1 ;	

} ;

PBXRESULT CSMTP::Invoke(IPB_Session * session, pbobject obj, pbmethodID mid, PBCallInfo * ci)
{

	Session = session ;

	switch( mid )
	{
	case SEND:
		{
		int	rc ;
		rc = Send ( ) ; 
		ci->returnValue->SetInt((pbint)rc );
		return PBX_OK;
		}
	case SETMESSAGE:
		{
		MessageArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbMessage = MessageArg->GetString() ;
		Message = (LPTSTR)session->GetString(pbMessage) ;
		#ifdef _DEBUG
			MessageBox( NULL, Message, _T ( "Message" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETMESSAGEHTML:
		{
		MessageArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbMessage = MessageArg->GetString() ;
		Message = (LPTSTR)session->GetString(pbMessage) ;
		#ifdef _DEBUG
			MessageBox( NULL, Message, _T ( "Message" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		MessageArg = session->AcquireValue ( ci->pArgs->GetAt(1) ) ;
		HTMLBody = MessageArg->GetBool() ;
		return PBX_OK;
		}
	case SETRECIPIENTEMAIL:
		{
		RecipientEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbRecipientEmail = RecipientEmailArg->GetString() ;
		RecipientEmail = (LPTSTR)session->GetString(pbRecipientEmail) ;
		#ifdef _DEBUG
			MessageBox( NULL, RecipientEmail, _T ( "RecipientEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETCCRECIPIENTEMAIL:
		{
		CCRecipientEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbCCRecipientEmail = CCRecipientEmailArg->GetString() ;
		CCRecipientEmail = (LPTSTR)session->GetString(pbCCRecipientEmail) ;
		#ifdef _DEBUG
			MessageBox( NULL, CCRecipientEmail, _T ( "CCRecipientEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETBCCRECIPIENTEMAIL:
		{
		BCCRecipientEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbBCCRecipientEmail = BCCRecipientEmailArg->GetString() ;
		BCCRecipientEmail = (LPTSTR)session->GetString(pbBCCRecipientEmail) ;
		#ifdef _DEBUG
			MessageBox( NULL, BCCRecipientEmail, _T ( "BCCRecipientEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETSENDEREMAIL:
		{
		SenderEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbSenderEmail = SenderEmailArg->GetString() ;
		SenderEmail = (LPTSTR)session->GetString ( pbSenderEmail ) ;
		#ifdef _DEBUG
			MessageBox( NULL, SenderEmail, _T ( "SenderEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETSMTPSERVER:
		{
		SMTPServerArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbSMTPServer = SMTPServerArg->GetString() ;
		SMTPServer = (LPTSTR)session->GetString ( pbSMTPServer ) ;
		#ifdef _DEBUG
			MessageBox( NULL, SMTPServer, _T ( "SMTPServer" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETSUBJECT: 
		{
		SubjectArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbSubject = SubjectArg->GetString() ;
		Subject = (LPTSTR)session->GetString ( pbSubject ) ;
		#ifdef _DEBUG
			MessageBox( NULL, Subject, _T ( "Subject" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETATTACHMENT:
		{
		AttachmentArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbAttachment = AttachmentArg->GetString() ;
		Attachment = (LPTSTR)session->GetString ( pbAttachment ) ;
		#ifdef _DEBUG
			MessageBox( NULL, Attachment, _T ( "Attachment" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETERRORMESSAGESON:
		{
		ErrorMessageBoxesOn = TRUE; 
		return PBX_OK;		
		}
	case SETERRORMESSAGESOFF:
		{
		ErrorMessageBoxesOn = FALSE; 
		return PBX_OK;		
		}
	case SETMSGCHARSET:
		{
		CharSetArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbCharSet = CharSetArg->GetString() ;
		CharSet = (LPTSTR)session->GetString ( pbCharSet ) ;
		#ifdef _DEBUG
			MessageBox( NULL, CharSet, _T ( "CharSet" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETUSERNAMEPASSWORD:
		{
		UsernameArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbUsername = UsernameArg->GetString() ;
		Username = (LPTSTR)session->GetString ( pbUsername ) ;
		PasswordArg = session->AcquireValue ( ci->pArgs->GetAt(1) ) ;
		pbstring pbPassword = PasswordArg->GetString() ;
		Password = (LPTSTR)session->GetString ( pbPassword ) ;
		#ifdef _DEBUG
			MessageBox( NULL, Username, _T ( "Username" ), MB_ICONEXCLAMATION | MB_OK );
			MessageBox( NULL, Password, _T ( "Password" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETNTLMAUTHENTICATION:
		{
			NTLMAuthentication = TRUE; 
			return PBX_OK;
		}
	case SETPORT:
		{
		IPB_Value * PortArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		Port = (int)PortArg->GetInt();
		return PBX_OK;
		}
	case SETSSL:
		{
			SSL = TRUE; 
			return PBX_OK;
		}
	case SETTLS:
		{
			TLS = TRUE; 
			return PBX_OK;
		}
		return PBX_E_INVOKE_METHOD_AMBIGUOUS;
	}
}

void CSMTP::CleanUp()
{
	//If any of our local IPB_Value pointers are populated
	//we need to release them to avoid a memory leak
	if (MessageArg)
	{
		Session->ReleaseValue ( MessageArg ) ;
	}
	if (SenderEmailArg)
	{
		Session->ReleaseValue ( SenderEmailArg ) ;
	}
	if (RecipientEmailArg)
	{
		Session->ReleaseValue ( RecipientEmailArg ) ;
	}
	if (CCRecipientEmailArg)
	{
		Session->ReleaseValue ( CCRecipientEmailArg ) ;
	}
	if (BCCRecipientEmailArg)
	{
		Session->ReleaseValue ( BCCRecipientEmailArg ) ;
	}
	if (SMTPServerArg)
	{
		Session->ReleaseValue ( SMTPServerArg ) ;
	}
	if (SubjectArg)
	{
		Session->ReleaseValue ( SubjectArg ) ;
	}
	if (AttachmentArg)
	{
		Session->ReleaseValue ( AttachmentArg ) ;
	}
	if (CharSetArg)
	{
		Session->ReleaseValue ( CharSetArg ) ;
	}
}

void CSMTP::Destroy()
{

	//Do cleanup before exiting
	CleanUp();

	delete this ;

}
