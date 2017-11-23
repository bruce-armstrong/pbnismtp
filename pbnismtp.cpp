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
		_T ( "subroutine SetReplyToEmail ( string pbreplytoemail )\n" )
		_T ( "subroutine SetSenderEmail ( string pbsenderemail )\n" )
		_T ( "subroutine SetSMTPServer ( string pbsmtpserver )\n" )
		_T ( "subroutine SetSubject ( string pbsubject )\n" )
		_T ( "subroutine SetAttachment ( string pbattachment )\n" )
		_T ( "subroutine SetErrorMessagesOn ( )\n" )
		_T ( "subroutine SetErrorMessagesOff ( )\n" )
		_T ( "subroutine SetCharSet ( string pbcharset )\n" )
		_T ( "subroutine SetUsernamePassword ( string pbusername, string pbpassword )\n" )
		_T ( "subroutine SetPort ( integer port )\n" )
		_T ( "subroutine SetAuthMethod ( integer authmethod )\n" )
		_T ( "subroutine SetConnectionType ( integer connecttype )\n" )
		_T ( "function string GetLastErrorMessage ( )\n")
		_T ( "subroutine SetPriority ( integer pbpriority )\n" )
		_T ( "subroutine SetReadReceiptRequested ( boolean pbboolean )\n" )
		_T ( "subroutine SetOriginatorDeliveryReportRequested ( boolean pbboolean )\n" )
		_T ( "subroutine SetMessage ( string pbplainmessage, string pbhtmlmessage )\n")
		_T ( "subroutine SetAttachment ( string pbattachment, string pbcontentid, string pbcontenttype )\n")
		_T ( "subroutine SetAttachmentBase64 ( string pbattachment, string pbfilename, string pbcontentid, string pbcontenttype )\n")
		_T ( "subroutine SetMailerName ( string pbname )\n")
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
	SETREPLYTOEMAIL = 6,
	SETSENDEREMAIL = 7,
	SETSMTPSERVER = 8,
	SETSUBJECT = 9,
	SETATTACHMENT = 10,
	SETERRORMESSAGESON = 11,   
	SETERRORMESSAGESOFF = 12,
	SETMSGCHARSET = 13,
	SETUSERNAMEPASSWORD = 14,
	SETPORT = 15,
	SETAUTHMETHOD = 16,
	SETCONNECTIONTYPE = 17,
	GETLASTERRORMESSAGE = 18,
	SETPRIORITY = 19,
	SETREADRECEIPTREQUEST = 20,
	SETORIGINATORDELIVERYREPORTREQUESTED = 21,
	SETMESSAGEPLAINHTML = 22,
	SETATTACHEMENTEXTRA = 23,
	SETATTACHMENTBASE64 = 24,
	SETMAILERNAME = 25,
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
	ReplyToEmailArg = NULL;
	SMTPServerArg = NULL ;
	SubjectArg = NULL ;
	AttachmentArg = NULL ;
	CharSetArg = NULL ;
	UsernameArg = NULL ;
	PasswordArg = NULL ;
	PriorityArg = NULL ;
	ReadReceiptArg = NULL ;
	DeliverReportArg = NULL ;	 

	Message = NULL;
	HTMLMessage = NULL;
	SenderEmail = NULL;
	RecipientEmail = NULL;
	CCRecipientEmail = NULL;
	BCCRecipientEmail = NULL;
	ReplyToEmail = NULL;
	SMTPServer = NULL;
	Subject = NULL; 
	Attachment = NULL;
	Username = NULL ;
	Password = NULL ;
	CharSet = _T("iso-8859-1");

	ErrorMessageBoxesOn = FALSE;    /* initialize to off  */	

	HTMLBody = FALSE ;
	PLAINHTMLBody = FALSE;

	Port = 25 ;
	AuthMethod = 5 ; /* Default to Auto Detect */
	ConnectType = 0 ; /* Default to Plain Text */

	MailerName = _T("PBNISMTP 3.2");

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

	TCHAR	title[] = _T("Debug");
	CPJNSMTPConnection connection;
	CPJNSMTPMessage msg;
	BOOLEAN CloseGracefully = TRUE;
	INT	ReturnCode = UNDEFINEDERR;

	// Validate the passed in values first
	if (SenderEmail == NULL)
	{
		return INVALIDSENDEREMAIL;
	}

	if (RecipientEmail == NULL && CCRecipientEmail == NULL && BCCRecipientEmail == NULL)
	{
		return INVALIDRECIPIENTEMAIL;
	}

	if (SMTPServer == NULL)
	{
		return INVALIDSMTPSERVER;
	}

	if (Subject == NULL)
	{
		Subject = _T("");
	}

	if (Message == NULL)
	{
		Message = _T("");
	}

	try
	{
		// Set the Return Code if Exception is raised for the following calls
		ReturnCode = INVALIDRECIPIENTEMAIL;

		// Set the Charset
		msg.SetCharset ( CharSet ) ;

		// Set the Sender
		CPJNSMTPAddress address ( SenderEmail ) ;
		msg.m_From = address ;
		
		// Set the Recipient
		msg.ParseMultipleRecipients ( RecipientEmail, msg.m_To ) ;
		
		// Set the CC
		if ( CCRecipientEmail != NULL )
		{
			msg.ParseMultipleRecipients ( CCRecipientEmail, msg.m_CC ) ;
		}
		// Set the BCC
		if ( BCCRecipientEmail != NULL )
		{
			msg.ParseMultipleRecipients ( BCCRecipientEmail, msg.m_BCC ) ;
		}
		// Set the Reply To
		if ( ReplyToEmail != NULL )
		{
			msg.ParseMultipleRecipients( ReplyToEmail, msg.m_ReplyTo );
		}

		// Set the Return Code if Exception is raised for the following calls
		ReturnCode = MESSAGEBODYERROR;

		// Set the subject and message
		msg.m_sSubject = (CString)Subject ; 
		msg.m_sXMailer = MailerName;

		// Set Read Receipt header
		if ( READRECEIPT == TRUE )
		{
			#ifdef _DEBUG
						MessageBox(NULL, _T("Set Disposition-Notification-To Header"), title, MB_ICONEXCLAMATION | MB_OK);
			#endif
			CString dString;
			dString.Format(_T("Disposition-Notification-To: %s"), SenderEmail);
			msg.m_CustomHeaders.push_back((CPJNSMTPString)dString);

		}

		// Set Delivery report header
		if ( DELIVERYREPORT == TRUE )
		{
			#ifdef _DEBUG
						MessageBox(NULL, _T("Set Delivery Report Header"), title, MB_ICONEXCLAMATION | MB_OK);
			#endif
			// set develivery report header
			msg.m_dwDSN = msg.DSN_SUCCESS;
		}

		// Set Priority Header
		if (PRIORITY != NULL)
		{
			#ifdef _DEBUG
						MessageBox(NULL, _T("Set Priority Header"), title, MB_ICONEXCLAMATION | MB_OK);
			#endif
			
			msg.m_Priority = msg.NoPriority;

			if (PRIORITY == 3)
			{
				msg.m_Priority = msg.HighPriority;
			}
			else if (PRIORITY == 2)
			{
				msg.m_Priority = msg.NormalPriority;
			}
			else if (PRIORITY == 1)
			{
				msg.m_Priority = msg.LowPriority;
			}
		}

		if ( HTMLBody == TRUE )
		{
			// Only HTML Message
			msg.SetMime(TRUE) ;
			msg.AddHTMLBody(Message);
		}
		else if (PLAINHTMLBody == TRUE)
		{
			// multipart plain/html message
			CPJNSMTPBodyPart RootPart;
			CPJNSMTPBodyPart MessagePart;
			CPJNSMTPBodyPart TextPart;
			CPJNSMTPBodyPart HTMLPart;

			RootPart = msg.m_RootPart;

			msg.SetMime(TRUE);

			MessagePart.SetCharset(RootPart.GetCharset().c_str());
			MessagePart.SetText(_T(""));
			MessagePart.SetContentType(_T("multipart/alternative"));

			TextPart.SetCharset(RootPart.GetCharset().c_str());
			TextPart.SetText(Message);
			TextPart.SetContentType(_T("text/plain"));
			MessagePart.AddChildBodyPart(TextPart);

			HTMLPart.SetCharset(RootPart.GetCharset().c_str());
			HTMLPart.SetText(HTMLMessage);
			HTMLPart.SetContentType(_T("text/html"));
			MessagePart.AddChildBodyPart(HTMLPart);

			msg.AddBodyPart(MessagePart);
		}
		else
		{
			// Only PlainText Message
			msg.AddTextBody (Message) ;
		}

		// Add any attachment
		if ( Attachment != NULL )
		{
			msg.AddMultipleAttachments(Attachment);
		}

		// Set any attachement in array
		for (std::vector<AttachmentFile*>::size_type i = 0; i < m_AttachmentFile.size(); i++)
		{
			CPJNSMTPBodyPart	BodyPart;
			BodyPart.SetFilename(m_AttachmentFile[i].Filename);
			BodyPart.SetContentID(m_AttachmentFile[i].ContentID);
			BodyPart.SetContentType(m_AttachmentFile[i].ContentType);
			msg.AddBodyPart(BodyPart);
		}

		// Set Embedded Attachements
		for (std::vector<AttachmentBase64*>::size_type i = 0; i < m_AttachmentBase64.size(); i++)
		{
			CPJNSMTPBodyPart	imageBodyPart;
			imageBodyPart.SetRawBody(m_AttachmentBase64[i].Base64);
			imageBodyPart.SetAttachment(TRUE);
			imageBodyPart.SetContentID(m_AttachmentBase64[i].ContentID);
			imageBodyPart.SetContentType(m_AttachmentBase64[i].ContentType);
			imageBodyPart.SetTitle(m_AttachmentBase64[i].FileName);
			msg.AddBodyPart(imageBodyPart);
		}

		// Set the Return Code if Exception is raised for the following calls
		ReturnCode = SMTPCONNECTFAILED;

		// Connect to SMTP server
		#ifdef _DEBUG
			MessageBox(NULL, _T("Connecting to SMTP Server"), title, MB_ICONEXCLAMATION | MB_OK);
		#endif
		connection.Connect(SMTPServer, CPJNSMTPConnection::AuthenticationMethod(AuthMethod), Username, Password, Port, CPJNSMTPConnection::ConnectionType(ConnectType));

		// Set the Return Code if Exception is raised for the following calls
		ReturnCode = MESSAGESENDERROR;

		// Now send the message
		#ifdef _DEBUG
			MessageBox(NULL, _T("Sending Message"), title, MB_ICONEXCLAMATION | MB_OK);
		#endif
		connection.SendMessage ( msg ) ;
		
		// Everything was Successful!
		ReturnCode = SUCCESSFUL;
	}
	#ifdef CPJNSMTP_MFC_EXTENSIONS
		catch (CPJNSMTPException* pEx)
		{
			CloseGracefully = FALSE;
			LastErrorMsg = pEx->m_sLastResponse;
			if (LastErrorMsg == "")
			{
				LastErrorMsg = (CString) pEx->GetErrorMessage().operator LPCTSTR();
			}
			if (ErrorMessageBoxesOn == TRUE)
			{
				//Display the error
				CString sMsg;
				sMsg.Format(_T("Could not send message: Error:%08X\nDescription:%s"), pEx->m_hr, LastErrorMsg);
				MessageBox(NULL, sMsg, title, MB_ICONEXCLAMATION | MB_OK);
			}
			pEx->Delete();
		}
	#else
		catch (CPJNSMTPException& e)
		{
			CloseGracefully = FALSE;
			LastErrorMsg = e.m_sLastResponse.c_str();
			if (LastErrorMsg == "")
			{
				TCHAR szError[4096];
				szError[0] = _T('\0');
				e.GetErrorMessage(szError, _countof(szError));
				LastErrorMsg = (CString) szError;
			}
			if (ErrorMessageBoxesOn == TRUE)
			{
				//Display the error
				CString sMsg;
				sMsg.Format(_T("Could not send message: Error:%08X\nDescription:%s"), e.m_hr, LastErrorMsg); // e.m_sLastResponse.c_str());
				MessageBox(NULL, sMsg, title, MB_ICONEXCLAMATION | MB_OK);
			}
		}
	#endif

	// Disconnect from the server
	try
	{
		if (connection.IsConnected())
			connection.Disconnect(CloseGracefully);
	}
	#ifdef CPJNSMTP_MFC_EXTENSIONS
		catch (CPJNSMTPException* pEx)
		{
			pEx->Delete();
		}
	#else
		catch (CPJNSMTPException& /*e*/)
		{
		}
	#endif
		
	return ReturnCode ;

} ;

PBXRESULT CSMTP::Invoke(IPB_Session * session, pbobject obj, pbmethodID mid, PBCallInfo * ci)
{

	Session = session ;

	switch( mid )
	{
	case SEND:
		{
		int	rc;
		rc = Send();
		ci->returnValue->SetInt((pbint)rc);
		return PBX_OK;
		}
	case SETMESSAGE:
		{
		MessageArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbMessage = MessageArg->GetString() ;
		Message = session->GetString(pbMessage) ;
		#ifdef _DEBUG
			MessageBox( NULL, Message, _T ( "Message" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETMESSAGEHTML:
		{
		MessageArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbMessage = MessageArg->GetString() ;
		Message = session->GetString(pbMessage) ;
		IPB_Value * HTMLArg = session->AcquireValue ( ci->pArgs->GetAt(1) ) ;
		HTMLBody = (bool)HTMLArg->GetBool() ;
		if (HTMLArg)
		{
			Session->ReleaseValue(HTMLArg);
		}
		#ifdef _DEBUG
			MessageBox(NULL, Message, _T("MessageHTML"), MB_ICONEXCLAMATION | MB_OK);
		#endif
		return PBX_OK;
		}
	case SETRECIPIENTEMAIL:
		{
		RecipientEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbRecipientEmail = RecipientEmailArg->GetString() ;
		RecipientEmail = session->GetString(pbRecipientEmail) ;
		#ifdef _DEBUG
			MessageBox( NULL, RecipientEmail, _T ( "RecipientEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETCCRECIPIENTEMAIL:
		{
		CCRecipientEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbCCRecipientEmail = CCRecipientEmailArg->GetString() ;
		CCRecipientEmail = session->GetString( pbCCRecipientEmail ) ;
		#ifdef _DEBUG
			MessageBox( NULL, CCRecipientEmail, _T ( "CCRecipientEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETBCCRECIPIENTEMAIL:
		{
		BCCRecipientEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbBCCRecipientEmail = BCCRecipientEmailArg->GetString() ;
		BCCRecipientEmail = session->GetString( pbBCCRecipientEmail ) ;
		#ifdef _DEBUG
			MessageBox( NULL, BCCRecipientEmail, _T ( "BCCRecipientEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETREPLYTOEMAIL:
		{
		ReplyToEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) );
		pbstring pbReplyToEmail = ReplyToEmailArg->GetString();
		ReplyToEmail = session->GetString( pbReplyToEmail );
		#ifdef _DEBUG
			MessageBox(NULL, ReplyToEmail, _T("ReplyToEmail"), MB_ICONEXCLAMATION | MB_OK);
		#endif
		return PBX_OK;
		}
	case SETSENDEREMAIL:
		{
		SenderEmailArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbSenderEmail = SenderEmailArg->GetString() ;
		SenderEmail = session->GetString ( pbSenderEmail ) ;
		#ifdef _DEBUG
			MessageBox( NULL, SenderEmail, _T ( "SenderEmail" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETSMTPSERVER:
		{
		SMTPServerArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbSMTPServer = SMTPServerArg->GetString() ;
		SMTPServer = session->GetString ( pbSMTPServer ) ;
		#ifdef _DEBUG
			MessageBox( NULL, SMTPServer, _T ( "SMTPServer" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETSUBJECT: 
		{
		SubjectArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbSubject = SubjectArg->GetString() ;
		Subject = session->GetString ( pbSubject ) ;
		#ifdef _DEBUG
			MessageBox( NULL, Subject, _T ( "Subject" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETATTACHMENT:
		{
		AttachmentArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbAttachment = AttachmentArg->GetString() ;
		Attachment = session->GetString ( pbAttachment ) ;
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
		CharSet = session->GetString ( pbCharSet ) ;
		#ifdef _DEBUG
			MessageBox( NULL, CharSet, _T ( "CharSet" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETUSERNAMEPASSWORD:
		{
		UsernameArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		pbstring pbUsername = UsernameArg->GetString() ;
		Username = session->GetString ( pbUsername ) ;
		PasswordArg = session->AcquireValue ( ci->pArgs->GetAt(1) ) ;
		pbstring pbPassword = PasswordArg->GetString() ;
		Password = session->GetString ( pbPassword ) ;
		#ifdef _DEBUG
			MessageBox( NULL, Username, _T ( "Username" ), MB_ICONEXCLAMATION | MB_OK );
		#endif
		return PBX_OK;
		}
	case SETPORT:
		{
		IPB_Value * PortArg = session->AcquireValue ( ci->pArgs->GetAt(0) ) ;
		Port = (int)PortArg->GetInt();
		if (PortArg)
		{
			Session->ReleaseValue(PortArg);
		}
		return PBX_OK;
		}
	case SETAUTHMETHOD:
		{
		IPB_Value * AuthArg = session->AcquireValue ( ci->pArgs->GetAt(0) );
		AuthMethod = (int)AuthArg->GetInt();
		if (AuthArg)
		{
			Session->ReleaseValue( AuthArg );
		}
		return PBX_OK;
		}
	case SETCONNECTIONTYPE:
		{
		IPB_Value * ConnectionArg = session->AcquireValue ( ci->pArgs->GetAt(0) );
		ConnectType = (int)ConnectionArg->GetInt();
		if (ConnectionArg)
		{
			Session->ReleaseValue( ConnectionArg );
		}
		return PBX_OK;
		}
	case GETLASTERRORMESSAGE:
		{
		ci->returnValue->SetString(LastErrorMsg);
		return PBX_OK;
		}
	case SETPRIORITY:
		{
			IPB_Value * PriorityArg = session->AcquireValue(ci->pArgs->GetAt(0));
			PRIORITY = (int)PriorityArg->GetInt();

			return PBX_OK;

		}
	case SETREADRECEIPTREQUEST:
		{
			IPB_Value * ReadReceiptArg = session->AcquireValue(ci->pArgs->GetAt(0));
			READRECEIPT = (bool)ReadReceiptArg->GetBool();
			if (ReadReceiptArg)
			{
				Session->ReleaseValue(ReadReceiptArg);
			}
			return PBX_OK;
		}
	case SETORIGINATORDELIVERYREPORTREQUESTED:
		{
			IPB_Value * DeliverReportArg = session->AcquireValue(ci->pArgs->GetAt(0));
			DELIVERYREPORT = (bool)DeliverReportArg->GetBool();
			if (DeliverReportArg)
			{
				Session->ReleaseValue(DeliverReportArg);
			}
			return PBX_OK;
		}
	case SETMESSAGEPLAINHTML:
		{
			// Plaintext message
			MessageArg = session->AcquireValue(ci->pArgs->GetAt(0));
			pbstring pbMessage = MessageArg->GetString();
			Message = session->GetString(pbMessage);

			// HTML Mesage
			MessageArg = session->AcquireValue(ci->pArgs->GetAt(1));
			pbstring pbHTMLMessage = MessageArg->GetString();
			HTMLMessage = session->GetString(pbHTMLMessage);

			PLAINHTMLBody = TRUE;

			#ifdef _DEBUG
					MessageBox(NULL, Message, _T("Message"), MB_ICONEXCLAMATION | MB_OK);
					MessageBox(NULL, HTMLMessage, _T("HTMLMessage"), MB_ICONEXCLAMATION | MB_OK);
			#endif
			return PBX_OK;
		}
	case SETATTACHEMENTEXTRA:
		{
			AttachmentArg = session->AcquireValue(ci->pArgs->GetAt(0));
			pbstring pbFileName = AttachmentArg->GetString();
			LPTSTR FileName = (LPTSTR)session->GetString(pbFileName);

			AttachmentArg = session->AcquireValue(ci->pArgs->GetAt(1));
			pbstring pbContentID = AttachmentArg->GetString();
			LPTSTR ContentID = (LPTSTR)session->GetString(pbContentID);

			AttachmentArg = session->AcquireValue(ci->pArgs->GetAt(2));
			pbstring pbContentType = AttachmentArg->GetString();
			LPTSTR ContentType = (LPTSTR)session->GetString(pbContentType);

			AttachmentFile	aFile;
			aFile.Filename = FileName;
			aFile.ContentID = ContentID;
			aFile.ContentType = ContentType;
			m_AttachmentFile.push_back(aFile);

			return PBX_OK;
		}
	case SETATTACHMENTBASE64:
		{
			AttachmentArg = session->AcquireValue(ci->pArgs->GetAt(0));
			pbstring pbBase64 = AttachmentArg->GetString();
			LPCSTR Base64 = (LPCSTR)session->GetString(pbBase64);

			AttachmentArg = session->AcquireValue(ci->pArgs->GetAt(1));
			pbstring pbFileName = AttachmentArg->GetString();
			LPTSTR FileName = (LPTSTR)session->GetString(pbFileName);

			AttachmentArg = session->AcquireValue(ci->pArgs->GetAt(2));
			pbstring pbContentID = AttachmentArg->GetString();
			LPTSTR ContentID = (LPTSTR)session->GetString(pbContentID);

			AttachmentArg = session->AcquireValue(ci->pArgs->GetAt(3));
			pbstring pbContentType = AttachmentArg->GetString();
			LPTSTR ContentType = (LPTSTR)session->GetString(pbContentType);

			AttachmentBase64	aBase64;
			aBase64.Base64 = Base64;
			aBase64.FileName = FileName;
			aBase64.ContentID = ContentID;
			aBase64.ContentType = ContentType;
			m_AttachmentBase64.push_back(aBase64);

			return PBX_OK;
		}
	case SETMAILERNAME:
		{
			MailerNameArg = session->AcquireValue(ci->pArgs->GetAt(0));
			pbstring pbMailerName = MailerNameArg->GetString();
			MailerName = session->GetString(pbMailerName);
			#ifdef _DEBUG
					MessageBox(NULL, MailerName, _T("MailerName"), MB_ICONEXCLAMATION | MB_OK);
			#endif
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
	if (ReplyToEmailArg)
	{
		Session->ReleaseValue( ReplyToEmailArg );
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
	if (UsernameArg)
	{
		Session->ReleaseValue( UsernameArg );
	}
	if (PasswordArg)
	{
		Session->ReleaseValue( PasswordArg );
	}
	if (ReadReceiptArg)
	{
		Session->ReleaseValue( ReadReceiptArg );
	}
	if (DeliverReportArg)
	{
		Session->ReleaseValue( DeliverReportArg );
	}
	if (PriorityArg)
	{
		Session->ReleaseValue( PriorityArg );
	}
}

void CSMTP::Destroy()
{

	//Do cleanup before exiting
	CleanUp();

	delete this ;

}
