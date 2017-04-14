#ifndef PBNISMTP_H
#define PBNISMTP_H

#pragma warning(disable:4715 4800)

#define INVALIDSENDEREMAIL -1
#define INVALIDRECIPIENTEMAIL -2
#define INVALIDSMTPSERVER -3
#define SMTPCONNECTFAILED -4
#define MESSAGESENDERROR -5
#define UNDEFINEDERR -99
#define FALSE 0
#define TRUE 1

class CSMTP : public IPBX_NonVisualObject
{
public:
	CSMTP();
	virtual ~CSMTP();

	PBXRESULT Invoke(
		IPB_Session * session, 
		pbobject obj, 
		pbmethodID mid, 
		PBCallInfo * ci
	);

protected:

	int Send();

	void CleanUp();

	LPTSTR Message; 
	LPTSTR SenderEmail;
	LPTSTR RecipientEmail;
	LPTSTR CCRecipientEmail;
	LPTSTR BCCRecipientEmail;
	LPTSTR SMTPServer ;
	LPTSTR Subject; 
	LPTSTR Attachment ;
	LPTSTR CharSet ;
	LPTSTR Username ;
	LPTSTR Password ;

	IPB_Value* MessageArg ;
	IPB_Value* SenderEmailArg ;
	IPB_Value* RecipientEmailArg ;
	IPB_Value* CCRecipientEmailArg ;
	IPB_Value* BCCRecipientEmailArg ;
	IPB_Value* SMTPServerArg ;
	IPB_Value* SubjectArg ;
	IPB_Value* AttachmentArg ;
	IPB_Value* CharSetArg ;
	IPB_Value* UsernameArg ;
	IPB_Value* PasswordArg ;

	IPB_Session*	Session ;

	int ErrorMessageBoxesOn ;

	bool HTMLBody ;

	bool NTLMAuthentication;
	int Port ;
	bool SSL ;
	bool TLS ;


private:

    static DWORD m_dwRefCount;

	void Destroy() ;

};

#endif 
