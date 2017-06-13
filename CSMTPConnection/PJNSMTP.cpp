/*
Module : PJNSmtp.cpp
Purpose: Implementation for a MFC class encapsulation of the SMTP protocol
Created: PJN / 22-05-1998
History: PJN / 15-06-1998 1. Fixed the case where a single dot occurs on its own
                          in the body of a message
                          2. Class now supports Reply-To Header Field
                          3. Class now supports file attachments
         PJN / 18-06-1998 1. Fixed a memory overwrite problem which was occurring 
                         with the buffer used for encoding base64 attachments
         PJN / 27-06-1998 1. The case where a line begins with a "." but contains
                          other text is now also catered for. See RFC821, Section 4.5.2
                          for further details.
                          2. m_sBody in CPJNSMTPMessage has now been made protected.
                          Client applications now should call AddBody instead. This
                          ensures that FixSingleDot is only called once even if the 
                          same message is sent a number of times.
                          3. Fixed a number of problems with how the MIME boundaries
                          were defined and sent.
                          4. Got rid of an unreferenced formal parameter 
                          compiler warning when doing a release build
         PJN / 11-09-1998 1. VC 5 project file is now provided
                          2. Attachment array which the message class contains now uses
                          references instead of pointers.
                          3. Now uses Sleep(0) to yield our time slice instead of Sleep(100),
                          this is the preferred way of writting polling style code in Win32
                          without serverly impacting performance.
                          4. All Trace statements now display the value as returned from
                          GetLastError
                          5. A number of extra asserts have been added
                          6. A AddMultipleRecipients function has been added which supports added a 
                          number of recipients at one time from a single string
                          7. Extra trace statements have been added to help in debugging
         PJN / 12-09-1998 1. Removed a couple of unreferenced variable compiler warnings when code
                          was compiled with Visual C++ 6.0
                          2. Fixed a major bug which was causing an ASSERT when the CSMTPAttachment
                          destructor was being called in the InitInstance of the sample app. 
                          This was inadvertingly introduced for the 1.2 release. The fix is to revert 
                          fix 2. as done on 11-09-1998. This will also help to reduce the number of 
                          attachment images kept in memory at one time.
         PJN / 18-01-1999 1. Full CC & BCC support has been added to the classes
         PJN / 22-02-1999 1. Addition of a Get and SetTitle function which allows a files attachment 
                          title to be different that the original filename
                          2. AddMultipleRecipients now ignores addresses if they are empty.
                          3. Improved the reading of responses back from the server by implementing
                          a growable receive buffer
                          4. timeout is now 60 seconds when building for debug
         PJN / 25-03-1999 1. Now sleeps for 250 ms instead of yielding the time slice. This helps 
                          reduce CPU usage when waiting for data to arrive in the socket
         PJN / 14-05-1999 1. Fixed a bug with the way the code generates time zone fields in the Date headers.
         PJN / 10-09-1999 1. Improved CPJNSMTPMessage::GetHeader to include mime field even when no attachments
                          are included.
         PJN / 16-02-2000 1. Fixed a problem which was occuring when code was compiled with VC++ 6.0.
         PJN / 19-03-2000 1. Fixed a problem in GetHeader on Non-English Windows machines
                          2. Now ships with a VC 5 workspace. I accidentally shipped a VC 6 version in one of the previous versions of the code.
                          3. Fixed a number of UNICODE problems
                          4. Updated the sample app to deliberately assert before connecting to the author's SMTP server.
         PJN / 28-03-2000 1. Set the release mode timeout to be 10 seconds. 2 seconds was causing problems for slow dial
                          up networking connections.
         PJN / 07-05-2000 1. Addition of some ASSERT's in CPJNSMTPSocket::Connect
         PP  / 16-06-2000 The following modifications were done by Puneet Pawaia
                          1. Removed the base64 encoder from this file
                          2. Added the base64 encoder/decoder implementation in a separate 
                          file. This was done because base64 decoding was not part of 
                          the previous implementation
                          3. Added support for ESMTP connection. The class now attempts to 
                          authenticate the user on the ESMTP server using the username and
                          passwords supplied. For this connect now takes the username and 
                          passwords as parameters. These can be null in which case ESMTP 
                          authentication is not attempted
                          4. This class can now handle AUTH LOGIN and AUTH LOGIN PLAIN authentication
                          schemes on 
         PP  / 19-06-2000 The following modifications were done by Puneet Pawaia
                          1. Added the files md5.* containing the MD5 digest generation code
                          after modifications so that it compiles with VC++ 6
                          2. Added the CRAM-MD5 login procedure.
         PJN / 10-07-2000 1. Fixed a problem with sending attachments > 1K in size
                          2. Changed the parameters to CPJNSMTPConnection::Connect
         PJN / 30-07-2000 1. Fixed a bug in AuthLogin which was transmitting the username and password
                          with an extra "=" which was causing the login to failure. Thanks to Victor Vogelpoel for
                          finding this.
         PJN / 05-09-2000 1. Added a CSMTP_NORSA preprocessor macro to allow the CPJNSMTPConnection code to be compiled
                          without the dependence on the RSA code.
         PJN / 28-12-2000 1. Removed an unused variable from ConnectESMTP.
                          2. Allowed the hostname as sent in the HELO command to be specified at run time 
                          in addition to using the hostname of the client machine
                          3. Fixed a problem where high ascii characters were not being properly encoded in
                          the quoted-printable version of the body sent.
                          4. Added support for user definable charset's for the message body.
                          5. Mime boundaries are now always sent irrespective if whether attachments are included or
                          not. This is required as the body is using quoted-printable.
                          6. Fixed a bug in sendLines which was causing small message bodies to be sent incorrectly
                          7. Now fully supports custom headers in the SMTP message
                          8. Fixed a copy and paste bug where the default port for the SMTP socket class was 110.
                          9. You can now specify the address on which the socket is bound. This enables the programmer
                          to decide on which NIC data should be sent from. This is especially useful on a machine
                          with multiple IP addresses.
                          10. Addition of functions in the SMTP connection class to auto dial and auto disconnect to 
                          the Internet if you so desire.
                          11. Sample app has been improved to allow Auto Dial and binding to IP addresses to be configured.
         PJN / 26-02-2001 1. Charset now defaults to ISO 8859-1 instead of us-ascii
                          2. Removed a number of unreferrenced variables from the sample app.
                          3. Headers are now encoded if they contain non ascii characters.
                          4. Fixed a bug in getLine, Thanks to Lev Evert for spotting this one.
                          5. Made the charset value a member of the message class instead of the connection class
                          6. Sample app now fully supports specifying the charset of the message
                          7. Added a AddMultipleAttachments method to CPJNSMTPMessage
                          8. Attachments can now be copied to each other via new methods in CSMTPAttachment
                          9. Message class now contains copies of the attachments instead of pointers to them
                          10. Sample app now allows multiple attachments to be added
                          11. Removed an unnecessary assert in QuotedPrintableEncode
                          12. Added a Mime flag to the CPJNSMTPMessage class which allows you to decide whether or not a message 
                          should be sent as MIME. Note that if you have attachments, then mime is assumed.
                          13. CSMTPAttachment class has now become CPJNSMTPBodyPart in anticipation of full support for MIME and 
                          MHTML email support
                          14. Updated copright message in source code and documentation
                          15. Fixed a bug in GetHeader related to _tsetlocale usage. Thanks to Sean McKinnon for spotting this
                          problem.
                          16. Fixed a bug in SendLines when sending small attachments. Thanks to Deng Tao for spotting this
                          problem.
                          17. Removed the need for SendLines function entirely.
                          18. Now fully supports HTML email (aka MHTML)
                          19. Updated copyright messages in code and in documentation
         PJN / 17-06-2001 1. Fixed a bug in CPJNSMTPMessage::HeaderEncode where spaces were not being interpreted correctly. Thanks
                          to Jim Alberico for spotting this.
                          2. Fixed 2 issues with ReadResponse both having to do with multi-line responses. Thanks to Chris Hanson 
                          for this update.
         PJN / 25-06-2001 1. Code now links in Winsock and RPCRT40 automatically. This avoids client code having to specify it in 
                          their linker settings. Thanks to Malte and Phillip for spotting this issue.
                          2. Updated sample code in documentation. Thanks to Phillip for spotting this.
                          3. Improved the code in CPJNSMTPBodyPart::SetText to ensure lines are correctly wrapped. Thanks to 
                          Thomas Moser for this fix.
         PJN / 01-07-2001 1. Modified QuotedPrintableEncode to prevent the code to enter in an infinite loop due to a long word i.e. 
                          bigger than SMTP_MAXLINE, in this case, the word is breaked. Thanks to Manuel Gustavo Saraiva for this fix.
                          2. Provided a new protected variable in CPJNSMTPBodyPart called m_bQuotedPrintable to bypass the 
                          QuotedPrintableEncode function in cases that we don't want that kind of correction. Again thanks to 
                          Manuel Gustavo Saraiva for this fix.
         PJN / 15-07-2001 1. Improved the error handling in the function CPJNSMTPMessage::AddMultipleAttachments. In addition the 
                          return value has been changed from BOOL to int
         PJN / 11-08-2001 1. Fixed a bug in QuotedPrintableEncode which was wrapping encoding characters across multiple lines. 
                          Thanks to Roy He for spotting this.
                          2. Provided a "SendMessage" method which sends a email directly from disk. This allows you 
                          to construct your own emails and the use the class just to do the sending. This function also has the 
                          advantage that it efficiently uses memory and reports progress.
                          3. Provided support for progress notification and cancelling via the "OnSendProgress" virtual method.
         PJN / 29-09-2001 1. Fixed a bug in ReadResponse which occured when you disconnected via Dial-Up Networking
                          while a connection was active. This was originally spotted in my POP3 class.
                          2. Fixed a problem in CPJNSMTPBodyPart::GetHeader which was always including the "quoted-printable" even when 
                          m_bQuotedPrintable for that body part was FALSE. Thanks to "jason" for spotting this.
         PJN / 12-10-2001 1. Fixed a problem where GetBody was reporting the size as 1 bigger than it should have been. Thanks
                          to "c f" for spotting this problem.
                          2. Fixed a bug in the TRACE statements when a socket connection cannot be made.
                          3. The sample app now displays a filter of "All Files" when selecting attachments to send
                          4. Fixed a problem sending 0 byte attachments. Thanks to Deng Tao for spotting this problem.
         PJN / 11-01-2002 1. Now includes a method to send a message directly from memory. Thanks to Tom Allebrandi for this
                          suggestion.
                          2. Change function name "IsReadible" to be "IsReadable". I was never very good at English!.
                          3. Fixed a bug in CPJNSMTPBodyPart::QuotedPrintableEncode. If a line was exactly 76 characters long plus 
                          \r\n it produced an invalid soft linebreak of "\r=\r\n\n". Thanks to Gerald Egert for spotting this problem.
         PJN / 29-07-2002 1. Fixed an access violation in CPJNSMTPBodyPart::QuotedPrintableEncode. Thanks to Fergus Gallagher for 
                          spotting this problem.
                          2. Fixed a problem where in very rare cases, the QuotedPrintableEncode function produces a single dot in a 
                          line, when inserting the "=" to avoid the mail server's maxline limit. I inserted a call to FixSingleDot 
                          after calling the QuotedPrintableEncode function in GetBody. Thanks to Andreas Kappler for spotting this
                          problem.
                          3. Fixed an issue in CPJNSMTPBodyPart::GetHeader where to ensure some mail clients can correctly handle
                          body parts and attachments which have a filename with spaces in it. Thanks to Andreas kappler for spotting
                          this problem.
                          4. QP encoding is now only used when you specify MIME. This fixes a bug as reported by David Terracino.
                          5. Removed now unused "API Reference" link in HTML file supported the code.
         PJN / 10-08-2002 1. Fixed a number of uncaught file exceptions in CPJNSMTPBodyPart::GetBody, CPJNSMTPMessage::SaveToDisk, and 
                          CPJNSMTPConnection::SendMessage. Thanks to John Allan Miller for reporting this problem.
                          2. The CPJNSMTPConnection::SendMessage version of the method which sends a file directly from disk, now fails if the
                          file is empty.
                          3. Improved the sample app by displaying a wait cursor while a message from file is being sent
                          4. Improved the sample app by warning the user if mail settings are missing and then bringing up the configuration
                          dialog.
         PJN / 20-09-2002 1. Fixed a problem where the code "Coder.EncodedMessage" was not being converted from an ASCII string to a UNICODE 
                          string in calls to CString::Format. This was occurring when sending the username and password for "AUTH LOGIN" support
                          in addition to sending the "username digest" for "AUTH CRAM-MD5" support. Thanks to Serhiy Pavlov for reporting this
                          problem.
                          2. Removed a number of calls to CString::Format and instead replaced with string literal CString constructors instead.
         PJN / 03-10-2002 1. Quoted printable encoding didn't work properly in UNICODE. (invalid conversion from TCHAR to BYTE). Thanks to
                          Serhiy Pavlov for reporting this problem.
                          2. Subject encoding didn't work properly in UNICODE. (invalid conversion from TCHAR to BYTE). Thanks to Serhiy Pavlov
                          for reporting this problem.
                          3. It didn't insert "charset" tag into root header for plain text messages (now it includes "charset" into plain text 
                          messages too). Thanks to Serhiy Pavlov for reporting this problem.
         PJN / 04-10-2002 1. Fixed an issue where the header / body separator was not being sent correctly for mails with attachments or when
                          the message is MIME encoded. Thanks to Serhiy Pavlov for reporting this problem.
                          2. Fixed a bug in QuotedPrintableEncode and HeaderEncode which was causing the errors. Thanks to Antonio Maiorano 
                          for reporting this problem.
                          3. Fixed an issue with an additional line separator being sent between the header and body of emails. This was only
                          evident in mail clients if a non mime email without attachments was sent.
         PJN / 11-12-2002 1. Review all TRACE statements for correctness
                          2. Provided a virtual OnError method which gets called with error information 
         PJN / 07-02-2003 1. Addition of a "BOOL bGracefully" argument to Disconnect so that when an application cancels the sending of a 
                          message you can pass FALSE and close the socket without properly terminating the SMTP conversation. Thanks to
                          "nabocorp" for this nice addition.
         PJN / 19-03-2003 1. Addition of copy constructors and operator= to CPJNSMTPMessage class. Thanks to Alexey Vasilyev for this suggestion.
         PJN / 13-04-2003 1. Fixed a bug in the handling of EHLO responses. Thanks to "Senior" for the bug report and the fix.
         PJN / 16-04-2003 1. Enhanced the CPJNSMTPAddress constructor to parse out the email address and friendly name rather than assume
                          it is an email address.
                          2. Reworked the syntax of the CPJNSMTPMessage::ParseMultipleRecipients method. Also now internally this function
                          uses the new CPJNSMTPAddress constructor
         PJN / 19-04-2003 1. Fixed a bug in the CPJNSMTPAddress constructor where I was mixing up the friendly name in the "<>" separators,
                          when it should have been the email address. 
         PJN / 04-05-2003 1. Fixed an issue where the class doesn't convert the mail body and subject to the wanted encoding but rather changes 
                          the encoding of the email in the email header. Since the issue of supporting several languages is a complicated one 
                          I've decided that we could settle for sending all CPJNSMTPConnection emails in UTF-8. I've added some conversion functions 
                          to the class that - at this point - always converts the body and subject of the email to UTF-8. A big thanks to Itamar 
                          Kerbel for this nice addition.
                          2. Moved code and sample app to VC 6.
         PJN / 05-05-2003 1. Reworked the way UTF8 encoding is now done. What you should do if you want to use UTF-8 encoding is set the charset
                          to "UTF-8" and use either QP or base 64 encoding.
                          2. Reworked the automatic encoding of the subject line to use the settings as taken from the root SMTP body part
                          3. Only the correct headers according to the MIME RFC are now encoded.
                          4. QP encoding is the encoding mechanism now always used for headers.
                          5. Headers are now only encoded if required to be encoded.
         PJN / 12-05-2003 1. Fixed a bug where the X-Mailer header was being sent incorrectly. Thanks to Paolo Vernazza for reporting this problem.
                          2. Addition of X-Mailer header is now optional. In addition sample app now does not send the X-Mailer header.
                          3. The sample app now does not send the Reply-To header.
         PJN / 18-08-2003 1. Modified the return value from the ConnectToInternet method. Instead of it being a boolean, it is now an enum. This
                          allows client code to differentiate between two conditions that it couldn't previously, namely when an internet connection
                          already existed or if a new connection (presumably a dial-up connection was made). This allows client code to then
                          disconnect if a new connection was required. Thanks to Pete Arends for this nice addition.
         PJN / 05-10-2003 1. Reworked the CPJNSMTPConnection::ReadResponse method to use the timeout provided by IsReadable rather than calling sleep.
                          Thanks to Clarke Brunt for reporting this issue.
         PJN / 03-11-2003 1. Simplified the code in CPJNSMTPConnection::ReadResponse. Thanks to Clarke Brunt for reporting this issue.
         PJN / 03-12-2003 1. Made code which checks the Login responses which contain "Username" and "Password" case insensitive. Thanks to 
                          Zhang xiao Pan for reporting this problem.
         PJN / 11-12-2003 1. Fixed an unreferrenced variable in CPJNSMTPConnection::OnError as reported by VC.Net 2003
                          2. DEBUG_NEW macro is now only used on VC 6 and not on VC 7 or VC 7.1. This avoids a problem on these compilers
                          where a conflict exists between it and the STL header files. Thanks to Alex Evans for reporting this problem.
         PJN / 31-01-2004 1. Fixed a bug in CPJNSMTPBodyPart::GetBody where the size of the body part was incorrectly calculating the size if the
                          encoded size was an exact multiple of 76. Thanks to Kurt Emanuelson and Enea Mansutti for reporting this problem.
         PJN / 07-02-2004 1. Fixed a bug in CPJNSMTPBodyPart::SetText where the code would enter an endless loop in the Replace function. It has now 
                          been replaced with CString::Replace. This now means that the class will not now compile on VC 5. Thanks to Johannes Philipp 
                          Grohs for reporting this problem.
                          2. Fixed a number of warnings when the code is compiled with VC.Net 2003. Thanks to Edward Livingston for reporting
                          this issue.
         PJN / 18-02-2004 1. You can now optionally set the priority of an email thro the variable CPJNSMTPMessage::m_Priority. Thanks to Edward 
                          Livingston for suggesting this addition.
         PJN / 04-03-2004 1. To avoid conflicts with the ATL Server class of the same name, namely "CSMTPConnection", the class is now called 
                          "CPJNSMTPConnection". To provide for easy upgrading of code, "CSMTPConnection" is now defined to be "CPJSMTPConnection" 
                          if the code detects that the ATL Server SMTP class is not included. Thanks to Ken Jones for reporting this issue.
         PJN / 13-03-2004 1. Fixed a problem where the CPJNSMTPBodyPart::m_dwMaxAttachmentSize value was not being copied in the CPJNSMTPBodyPart::operator=
                          method. Thanks to Gerald Egert for reporting this problem and the fix.
         PJN / 05-06-2004 1. Fixed a bug in CSMTPConnection::ReadResponse, where the wrong parameters were being null terminated. Thanks to "caowen" 
                          for this update.
                          2. Updated the CSMTPConnection::ReadResponse function to handle multiline responses in all cases. Thanks to Thomas Liebethal
                          for reporting this issue.
         PJN / 07-06-2004 1. Fixed a potential buffer overflow issue in CPJNSMTPConnection::ReadResponse when certain multi line responses are received
         PJN / 30-09-2004 1. Fixed a parsing issue in CPJNSMTPConnection::ReadResponse when multi line responses are read in as multiple packets. 
                          Thanks to Mark Smith for reporting this problem.
                          2. Reworked the code which supports the various authentication mechanisms to support the correct terms. What was called 
                          "AUTH LOGIN PLAIN" support now more correctly uses the term "AUTH PLAIN". The names of the functions and enums have now 
                          also been reworked.
                          3. Did a review of the sample app code provided with the class so that the name of the modules, projects, exe etc is now
                          "PJNSMTPApp". Also reworked the name of some helper classes as well as the module name which supports the main class.
                          4. Reworked CPJNSMTPConnection::AuthLoginPlain (now called AuthPlain) to correctly handle the case where an invalid
                          response is received when expecting the username: response. Thanks to Mark Smith for reporting this problem.
         PJN / 23-12-2004 "Name" field in Content-Type headers is now quoted just like the filename field. Thanks to Mark Smith for reporting this
                          issue in conjunction with the mail client Eudora.
         PJN / 23-01-2005 1. All classes now uses exceptions to indicate errors. This means the whole area of handling errors in the code a whole lot
                          simpler. For example the OnError mechanism is gone along with all the string literals in the code. The actual code itself
                          is a whole lot simpler also. You should carefully review your code as a lot of the return values from methods (especially
                          CPJNSMTPConnection) are now void and will throw CSMTPException's. Thanks to Mark Smith for prompting this update.
                          2. General tidy up of the code following the testing of the new exception based code.
         PJN / 06-04-2005 1. Addition of PJNSMTP_EXT_CLASS define to the classes to allow the classes to be easily incorporated into extension DLLs.
                          Thanks to Arnaud Faucher for suggesting this addition.
                          2. Now support NTLM authentication. Thanks to Arnaud Faucher for this nice addition. NTLM Authentication is provided by
                          a new reusable class called "CNTLMClientAuth" in PJNNTLMAuth.cpp/h".
                          3. Fixed a bug in the sample app in the persistence of the Authentication setting. Thanks to Arnaud Faucher for reporting
                          this issue.
         PJN / 26-03-2005 1. Fixed compile problems with the code when the Force Conformance In For Loop Scope compile setting is used in 
                          Visual Studio .NET 2003. Thanks to Alexey Kuznetsov for reporting this problem.
                          2. Fixed compile problems when the code is compiled using the Detect 64-bit Portability Issues setting in Visual Studio
                          .NET 2003.
         PJN / 18-04-2005 1. Addition of a simple IsConnected() function inline with the author's CPop3Connection class. Thanks to Alexey Kuznetsov
                          for prompting this addition.
                          2. Addition of a MXLookup function which provides for convenient DNS lookups of MX records (the IP addresses of SMTP mail 
                          servers for a specific host domain). Internally this new function uses the new DNS functions which are provided with 
                          Windows 2000 or later. To ensure that the code continues to work correctly on earlier versions of Windows, the function 
                          pointers for the required functions are constructed at runtime using GetProcAddress. You can use this function to 
                          discover the IP address of the mail servers responsible for a specific domain. This allows you to deliver email directly 
                          to a domain rather than through a third party mail server. Thanks to Hans Dietrich for suggesting this nice addition.
         PJN / 23-04-2005 1. The code now uses a MS Crypto API implementation of the MD5 HMAC algorithm instead of custom code based on the RSA 
                          MD5 code. This means that there is no need for the RSA MD5.cpp, MD5.h and glob-md5.h to be included in client applications. 
                          In addition the build configurations for excluding the RSA MD5 code has been removed from the sample app. The new classes to 
                          perform the MD5 hashing are contained in PJNMD5.h and are generic enough to be included in client applications in their
                          own right.
         PJN / 01-05-2005 1. Now uses the author'ss CWSocket sockets wrapper class.
                          2. Added support for connecting via Socks4, Socks5 and HTTP proxies 
                          3. The last parameter to Connect has now been removed and is configured via an accessor / mutator pair Get/SetBoundAddress.
                          4. CSMTPException class is now called CPJNSMTPException.
                          5. Class now uses the Base64 class as provided with CWSocket. This means that the modules Base64Coder.cpp/h are no longer
                          distributed or required and should be removed from any projects which use the updated SMTP classes.
                          6. Fixed a bug in SaveToDisk method whereby existing files were not being truncated prior to saving the email message.
                          7. All calls to allocate heap memory now generate a CPJNSMTPException if the allocation fails.
                          8. Addition of a MXLookupAvailable() function so that client code can know a priori if DNS MX lookups are supported.
                          9. Sample app now allows you to use the MXLookup method to determine the mail server to use based on the first email address
                          you provide in the To field.
         PJN / 03-05-2005 1. Fixed a number of warnings when the code is compiled in Visual Studio .NET 2003. Thanks to Nigel Delaforce for
                          reporting these issues.
         PJN / 24-07-2005 1. Now includes support for parsing "(...)" comments in addresses. Thanks to Alexey Kuznetsov for this nice addition.
                          2. Fixed an issue where the mail header could contain a "Content-Transfer-Encoding" header. This header should only be 
                          applied to headers of MIME body parts rather than the mail header itself. Thanks to Jeroen van der Laarse, Terence Dwyer
                          and Bostjan Erzen for reporting this issue and the required fix.
                          3. Now calling both AddTextBody and AddHTMLBody both set the root body part MIME type to multipart/mixed. Previously
                          the code comments said one thing and did another (set it to "multipart/related"). Thanks to Bostjan Erzen for reporting
                          this issue.
         PJN / 14-08-2005 1. Minor update to include comments next to various locations in the code to inform users what to do if they get compile
                          errors as a result of compiling the code without the Platform SDK being installed. Thanks to Gudjon Adalsteinsson for 
                          prompting this update.
                          2. Fixed an issue when the code calls the function DnsQuery when the code is compiled for UNICODE. This was the result of 
                          a documentation error in the Platform SDK which incorrectly says that DnsQuery_W takes a ASCII string when in actual fact 
                          it takes a UNICODE string. Thanks to John Oustalet III for reporting this issue.
         PJN / 05-09-2005 1. Fixed an issue in MXLookup method where memory was not being freed correctly in the call to DnsRecordListFree. Thanks 
                          to Raimund Mettendorf for reporting this bug.
                          2. Fixed an issue where lines which had a single dot i.e. "." did not get mailed correctly. Thanks to Raimund Mettendorf
                          for reporting this issue.
                          3. Removed the call to FixSingleDot from SetText method as it is done at runtime as the email as being sent. This was 
                          causing problems for fixing up the dot issue when MIME encoded messages were being sent.
                          4. Fixed an issue where the FixSingleDot function was being called when saving a message to disk. This was incorrect
                          and it should only be done when actually sending an SMTP email.
         PJN / 07-09-2005 1. Fixed another single dot issue, this time where the body is composed of just a single dot. Thanks to Raimund Mettendorf
                          for reporting this issue.
                          2. Fixed a typo in the accessor function CPJNSMTPConnection::GetBoundAddress which was incorrectly called 
                          CPJNSMTPConnection::SetBoundAddress.
         PJN / 29-09-2005 1. Removed linkage to secur32.lib as all SSPI functions are now constructed at runtime using GetProcAddress. Thanks to 
                          Emir Kapic for this update.
                          2. 250 or 251 is now considered a successful response to the RCPT command. Thanks to Jian Peng for reporting this issue.
                          3. Fixed potential memory leaks in the implementation of NTLM authentication as the code was throwing exceptions
                          which the authentication code was not expected to handle.
         PJN / 02-10-2005 1. MXLookup functionality of the class can now be excluded via the preprocessor define "CPJNSMTP_NOMXLOOKUP". Thanks to
                          Thiago Luiz for prompting this update. 
                          2. Sample app now does not bother asking for username and password if authentication type is NTLM. Thanks to Emir Kapic
                          for spotting this issue.
                          3. Documentation has been updated to refer to issues with NTLM Authentication and SSPI. Thanks to Emir Kapic for prompting
                          this update.
                          4. Fixed an update error done for the previous version which incorrectly supported the additional 251 command for the "MAIL"
                          command instead of more correctly the "RCPT" command. Thanks to Jian Peng for reporting this issue.
         PJN / 07-10-2005 1. Fixed an issue where file attachments were incorrectly including a charset value in their headers. Thanks to Asle Rokstad
                          for reporting this issue.
         PJN / 16-11-2005 1. Now includes full support for connecting via SSL. For example sending mail thro Gmail using the server 
                          "smtp.gmail.com" on port 465. Thanks to Vladimir Sukhoy for providing this very nice addition (using of course the author's 
                          very own OpenSSL wrappers!).
         PJN / 29-06-2006 1. Code now uses new C++ style casts rather than old style C casts where necessary. 
                          2. Integrated ThrowPJNSMTPException function into CPJNSMTPConnection class and renamed to ThrowPJNSMTPException
                          3. Combined the functionality of the _PJNSMTP_DATA class into the main CPJNSMTPConnection class.
                          4. Updated the code to clean compile on VC 2005
                          5. Updated copyright details.
                          6. Removed unnecessary checks for failure to allocate via new since in MFC these will throw CMemoryExceptions!.
                          7. Optimized CPJNSMTPException constructor code
         PJN / 14-07-2006 1. Fixed a bug in CPJNSMTPConnection where when compiled in UNICODE mode, NTLM authentication would fail. Thanks to 
                          Wouter Demuynck for providing the fix for this bug.
                          2. Reworked the CPJNSMTPConnection::NTLMAuthPhase(*) methods to throw standard CPJNSMTPException exceptions if there are any 
                          problems detected. This can be done now that the CNTLMClientAuth class is resilient to exceptions being thrown.
         PJN / 17-07-2006 1. Updated the code in the sample app to allow > 30000 characters to be entered into the edit box which contains the body
                          of the email. This is achieved by placing a call to CEdit::SetLimitText in the OnInitDialog method of the dialog. Thanks to 
                          Thomas Noone for reporting this issue.
         PJN / 13-10-2006 1. Fixed an issue in Disconnect where if the call to ReadCommandResponse threw an exception, we would not close down the 
                          socket or set m_bConnected to FALSE. The code has been updated to ensure these are now done. This just provides a more
                          consistent debugging experience. Thanks to Alexey Kuznetsov for reporting this issue.
                          2. CPJNSMTPException::GetErrorMessage now uses the user default locale for formatting error strings.
                          3. CPJNSMTPMessage::ParseMultipleRecipients now supports email addresses of the form ""Friendly, More Friendly" <emailaddress>"
                          Thanks to Xiao-li Ling for this very nice addition.
                          4. Fixed some issues with cleaning up of header and body parts in CPJNSMTPConnection::SendBodyPart and 
                          CPJNSMTPMessage::WriteToDisk. Thanks to Xiao-li Ling for reporting this issue.
         PJN / 09-11-2006 1. Reverted CPJNSMTPException::GetErrorMessage to use the system default locale. This is consistent with how MFC
                          does its own error handling.
                          2. Now includes comprehensive support for DSN's (Delivery Status Notifications) as specified in RFC 3461. Thanks to 
                          Riccardo Raccuglia for prompting this update.
                          3. CPJNSMTPBodyPart::GetBody and CPJNSMTPConnection::SendMessage(const CString& sMessageOnFile... now does not open the disk
                          file for exclusive read access, meaning that other apps can read from the file. Thanks to Wouter Demuynck for reporting
                          this issue.
                          4. Fixed an issue in CPJNSMTPConnection::SendMessage(const CString& sMessageOnFile... where under certain circumstances
                          we could end up deleting the local "pSendBuf" buffer twice. Thanks to Izidor Rozman for reporting this issue.
         PJN / 30-03-2007 1. Fixed a bug in CPJNSMTPConnection::SetHeloHostname where an unitialized stack variable could potentially be used. Thanks
                          to Anthony Kowalski for reporting this bug.
                          2. Updated copyright details.
         PJN / 01-08-2007 1. AuthCramMD5, ConnectESMTP, ConnectSMTP, AuthLogin, and AuthPlain methods have now been made virtual.
                          2. Now includes support for a Auto Authentication protocol. This will detect what authentication methods the SMTP server 
                          supports and uses the "most secure" method available. If you do not agree with the order in which the protocol is chosen,
                          a virtual function called "ChooseAuthenticationMethod" is provided. Thanks to "zhangbo" for providing this very nice 
                          addition to the code.
                          3. Removed the definition for the defunct "AuthNTLM" method.
         PJN / 13-11-2007 1. Major update to the code to use VC2005 specific features. This version of the code and onwards will be supported only
                          on VC 2005 or later. Thanks to Andrey Babushkin for prompting this update. Please note that there has been breaking changes
                          to the names and layout of various methods and member variables. You should carefully analyze the sample app included in the
                          download as well as your client code for the required updates you will need to make.
                          2. Sample app included in download is now build for Unicode and links dynamically to MFC. It also links dynamically to the 
                          latest OpenSSL 0.9.8g dlls.
         PJN / 19-11-2007 1. The "Date:" header is now formed using static lookup arrays for the day of week and month names instead of using 
                          CTime::Format. This helps avoid generating invalid Date headers on non English version of Windows where the local language 
                          versions of these strings would end up being used. Thanks to feedback at the codeproject page for the bug report (specifically 
                          http://www.codeproject.com/internet/csmtpconn.asp?df=100&forumid=472&mpp=50&select=2054493#xx2054493xx for reporting this issue).
         PJN / 22-12-2007 1. Following feedback from Selwyn Stevens that the AUTH PLAIN mechanism was failing for him while trying to send mail using the
                          gmail SMTP server, I decided to go back and re-research the area of SMTP authentication again<g>. It turns out that the way I was 
                          implementing AUTH PLAIN support was not the same as that specified in the relevant RFCs. The code has now been updated to be
                          compliant with that specified in the RFC. I've confirmed the fix as operating correctly by using the gmail SMTP mail server
                          at "smtp.gmail.com" on the SMTP/SSL port of 465. I've also taken the opportunity to update the documentation to include a section
                          with links to various RFC's and web pages for background reference material
         PJN / 24-12-2007 1. CPJNSMTPException::GetErrorMessage now uses the FORMAT_MESSAGE_IGNORE_INSERTS flag. For more information please see Raymond
                          Chen's blog at http://blogs.msdn.com/oldnewthing/archive/2007/11/28/6564257.aspx. Thanks to Alexey Kuznetsov for reporting this
                          issue.
                          2. All username and password temp strings are now securely destroyed using SecureZeroMemory.
         PJN / 31-12-2007 1. Updated the sample app to clean compile on VC 2008
                          2. Fixed a x64 compile problems in RemoveCustomHeader & AddMultipleAttachments.
                          3. Fixed a x64 compile problem in CPJNSMTPAppConfigurationDlg::CBAddStringAndData
                          4. CPJNSMTPException::GetErrorMessage now uses Checked::tcsncpy_s
                          5. Other Minor coding updates to CPJNSMTPException::GetErrorMessage
         PJN / 02-02-2008 1. Updated copyright details.
                          2. Fixed a bug in CPJNSMTPMessage::FormDateHeader where sometimes an invalid "Date:" header is created. The bug will manifest 
                          itself for anywhere that is ahead of GMT, where the difference isn't a multiple of one hour. According to the Windows Time Zone 
                          applet, this bug would have occured in: Tehran, Kabul, Chennai, Kolkata, Mumbai, New Delhi, Sri Jayawardenepura, Kathmandu, Yangon, 
                          Adelaide and Darwin. Thanks to Selwyn Stevens for reporting this bug and providing the fix.
         PJN / 01-03-2008 1. CPJNSMTP Priority enum values, specifically NO_PRIORITY have been renamed to avoid clashing with a #define of the same name in the
                          Windows SDK header file WinSpool.h. The enum values for DSN_RETURN_TYPE have also been renamed to maintain consistency. Thanks to 
                          Zoran Buntic for reporting this issue. 
                          2. Fixed compile problems related to the recent changes in the Base64 class. Thanks to Mat Berchtold for reporting this issue.
                          3. Since the code is now for VC 2005 or later only, the code now uses the Base64 encoding support from the ATL atlenc.h header file.
                          Thanks to Mat Berchtold for reporting this optimization. This means that client projects no longer need to include Base64.cpp/h in
                          their projects.
         PJN / 31-05-2008 1. Code now compiles cleanly using Code Analysis (/analyze)
                          2. Removed the use of the function QuotedPrintableEncode and replaced with ATL::QPEncode
                          3. Removed the use of the function QEncode and replaced with ATL::QEncode
                          4. Reworked ReadResponse to use CStringA in line with the implementation in the POP3 class of the author.
         PJN / 20-07-2008 1. Fixed a bug in ReadResponse where the code is determining if it has received the terminator. Thanks to Tony Cool for reporting
                          this bug.
         PJN / 27-07-2008 1. Updated code to compile correctly using _ATL_CSTRING_EXPLICIT_CONSTRUCTORS define
                          2. CPJNSMTPMessage::GetHeader now correctly ensures all long headers are properly folded. In addition this function has been reworked
                          to create the header internally as an ASCII string rather than as a TCHAR style CString.
         PJN / 16-08-2008 1. Updated the AUTH_AUTO login support to fall back to no authentication if no authentication scheme is supported by the SMTP server.
                          Thanks to Mat Berchtold for this update.
         PJN / 02-11-2008 1. Improvements to CPJNSMTPAddress constructor which takes a single string. The code now removes redundant quotes.
                          2. CPJNSMTPAddress::GetRegularFormat now forms the regular form of the email address before it Q encodes it
                          3. The sample app now uses DPAPI to encrypt the username & password configuration settings
                          4. The sample app is now linked against the latest OpenSSL v0.9.8i dlls
                          4. Updated coding references in the html documentation. In addition the zip file now includes the OpenSSL dlls. Thanks to Michael 
                          Grove for reporting this issues.
         PJN / 03-11-2008 1. For best compatibility purposes with existing SMTP mail servers, all message headers which include email addresses are now not
                          Q encoded. Also the friendly part of the email address is quoted. Thanks to Christian Egging for reporting this issue.
                          2. Optimized use of CT2A class throughout the code
         PJN / 25-06-2009 1. Updated copyright details
                          2. Fixed a bug where the Reply-To email address was not correctly added to the message. Thanks to Dmitriy Maksimov for reporting this
                          issue.
                          3. Updated the sample app's project settings to more modern default values.
                          4. Updated the sample exe to ship with OpenSSL v0.98k
                          5. Fixed a bug in CPJNSMTPConnection::SendBodyPart where it would cause errors from OpenSSL when trying to send a body of 0 bytes in size.
                          This can occur when MHTML based emails are sent and you are using a "multipart/related" body part. Thanks to "atota" for reporting this
                          issue
         PJN / 13-11-2009 1. Now includes comprehensive support for MDN's (Message Disposition Notifications) as specified in RFC 3798. Thanks to 
                          Ron Brunton for prompting this update.
                          2. Rewrote CPJNSMTPMessage::SaveToDisk method to use CAtlFile
         PJN / 17-12-2009 1. The sample app is now linked against the latest OpenSSL v0.9.8l dlls
                          2. Updated the sample code in the documentation to better describe how to use the various classes. Thanks to Dionisis Kofos for reporting
                          this issue.
         PJN / 23-05-2010 1. Updated copyright details
                          2. Updated sample app to compile cleanly on Visual Studio 2010
                          3. The sample app is now linked against the latest OpenSSL v1.0.0 dlls
                          4. Removed an unused "sRet" variable in CPJNSMTPConnection::AuthCramMD5. THanks to Alain Danteny for reporting this issue.
                          5. Replaced all calls to memcpy with memcpy_s
                          6. The code now supports STARTTLS encryption as defined in RFC 3207.
                          7. If the call to DnsQuery fails, the error value is now preserved using SetLastError
                          8. AddTextBody now sets the mime type of the root body part multipart/alternative. This is a more appropriate value to use which is better
                          supported by more email clients. Thanks to Thane Hubbell for reporting this issue.
         PJN / 04-06-2010 1. Updated the code and sample app to compile cleanly when CPJNSMTP_NOSSL is defined. Thanks to "loggerlogger" for reporting this issue.
         PJN / 10-07-2010 1. Following feedback from multiple users of PJNSMTP, including Chris Bamford and "loggerlogger", the change from "multipart/mixed" to
                          "multipart/alternative" has now been reverted. Using "multipart/mixed" is a more appropriate value to use. Testing by various clients
                          have shown that this setting works correctly when sending HTML based email in GMail, Thunderbird and Outlook. Unfortunately this is the 
                          price I must pay when accepting end user contributions to my code base which I cannot easily test. Going forward I plan to be much more
                          discerning on what modifications I will accept to avoid these issues.
                          2. The CPJNSMTPBodyPart now has the concept of whether or not the body is considered an attachment. Previously the code assumed that if 
                          you set the m_sFilename parameter that the body part was an attachment. Instead now the code uses the "m_bAttachment" member variable. 
                          This allows an in memory representation of an attachment to be added to a message without the need to write it to an intermediate file
                          first. For example, here would be the series of steps you would need to go through to add an in memory jpeg image to an email and have it
                          appear as an attachment:
                          
                          const BYTE* pInMemoryImage = pointer to your image;
                          DWORD dwMemoryImageSize = size in bytes of pInMemoryImage;
                          CPJNSMPTBase64Encode encode;
                          encode.Encode(pInMemoryImage, dwMemoryImageSize, ATL_BASE64_FLAG_NONE);
                          CPJNSMTPBodyPart imageBodyPart;
                          imageBodyPart.SetRawBody(encode.Result());
                          imageBodyPart.SetAttachment(TRUE);
                          imageBodyPart.SetContentType(_T("image/jpg"));
                          imageBodyPart.SetTitle(_T("the name of the image"));
                          CPJNSMTPMessage message;
                          message.AddBodyPart(imageBodyPart);
                          
                          Thanks to Stephan Eizinga for suggesting this nice addition.
         PJN / 28-11-2010 1. AddTextBody and AddHTMLBody now allow the root body part's MIME type to be changed. Thanks to Thane Hubbell for prompting this update.
                          2. Added a CPJNSMTPBodyPart::GetBoundary method.
                          3. Added some sample code to the sample app's CPJNSMTPAppDlg constructor to show how to create a message used SMTP body parts. Thanks to
                          Thane Hubbell for prompting this update.
                          4. CPJNSMTPBodyPart::GetHeader now allows quoted-printable and raw attachments to be used
                          5. The sample app is now linked against the latest OpenSSL v1.0.0b dlls
         PJN / 18-12-2010 1. Remove the methods Set/GetQuotedPrintable and Set/GetBase64 from CPJNSMTPBodyPart and replaced them with new 
                          Set/GetContentTransferEncoding methods which works with a simple enum. 
                          2. CPJNSMTPBodyPart::SetAttachment and CPJNSMTPBodyPart::SetFilename now automatically sets the Content-Transfer-Encoding to base64 for 
                          attachments. Thanks to Christian Egging for reporting this issue.
                          3. The sample app is now linked against the latest OpenSSL v1.0.0c dlls
         PJN / 08-02-2011 1. Updated copyright details
                          2. Updated code to support latest SSL and Sockets class from the author. This means that the code now supports IPv6 SMTP servers
                          3. Connect method now allows binding to a specific IP address
         PJN / 13-02-2011 1. Remove the IP binding address parameter from the Connect method as there was already support for binding via the Set/GetBoundAddress
                          methods.
                          2. Set/GetBoundAddress have been renamed Set/GetBindAddress for consistency with the sockets class
         PJN / 01-04-2011 1. Reintroduced the concept of Address Header encoding. By default this setting is off, but can be enabled via a new
                          CPJNSMPTMessage::m_bAddressHeaderEncoding boolean value. Thanks to Bostjan Erzen for reporting this issue.
         PJN / 04-12-2011 1. Updated CPJNSMTPBodyPart::GetHeader to encode the title of a filename if it contains non-ASCII characters. Thanks to Anders Gustafsson
                          for reporting this issue.
                          2. The sample app is now linked against the latst OpenSSL v1.0.0e dlls.
         PJN / 12-08-2012 1. STARTTLS support code now uses TLSv1_client_method OpenSSL function instead of SSLv23_client_method. This fixes a problem where sending
                          emails using Hotmail / Windows Live was failing.         
                          2. CPJNSMTPConnection::_Close now provides a bGracefully parameter
                          3. Updated the code to clean compile on VC 2012
                          4. The sample app is now linked against the latst OpenSSL v1.0.1c dlls.
                          5. SendMessage now throws a CPJNSMTPException if OnSendProgress returns FALSE.
                          6. Addition of a new ConnectionType called "AutoUpgradeToSTARTTLS" which will automatically upgrade a plain text connection to STARTTLS
                          if the code detects that the SMTP server supports STARTTLS. If the server doesn't support STARTTLS then the connection will remain as if
                          you specified "PlainText"
                          7. Reworked the code to determine if it should connect using EHLO instead of HELO into a new virtual method called DoEHLO.
                          8. Reworked the internals of the ConnectESMTP method.
         PJN / 23-09-2012 1. Removed SetContentBase and GetContentBase methods as Content-Base header is deprecated from MHTML. For details see 
                          http://www.ietf.org/mail-archive/web/apps-discuss/current/msg03220.html. Thanks to Mat Berchtold for reporting this issue.
                          2. Removed the now defunct parameter from the AddHTMLBody method.
                          3. Removed the FoldHeader method as it was only used by the m_sXMailer header.
                          4. Updated the logic in the FolderSubjectHeader method to correctly handle the last line of text to fold.
         PJN / 24-09-2012 1. Removed an unnecessary line of code from CPJNSMTPConnection::SendRCPTForRecipient. Thanks to Mat Berchtold for reporting this issue.
         PJN / 30-09-2012 1. Updated the code to avoid DLL planting security issues when calling LoadLibrary. Thanks to Mat Berchtold for reporting this issue.
         PJN / 25-11-2012 1. Fixed some issues in the code when CPJNSMTP_NOSSL was defined. Looks like support for compiling without SSL support has been 
                          broken for a few versions. Thanks to Bostjan Erzen for reporting this issue.
         PJN / 12-03-2013 1. Updated copyright details.
                          2. CPJNSMTPAddressArray typedef is now defined in the CPJNSMTPMessage class instead of in the global namespace. In addition this typedef
                          class is now known as "CAddressArray"
                          3. The sample project included in the download now links to the various OpenSSL libraries via defines in stdafx.h instead of via project
                          settings. Thanks to Ed Nafziger for suggesting this nice addition. 
                          4. The sample app no longer adds the author's hotmail address to the list of MDN's. Thanks to Ed Nafziger for spotting this.
                          5. The sample app is now linked against the latest OpenSSL v1.0.1e dlls.
                          6. Sample app now statically links to MFC
         PJN / 03-06-2013 1. The demo app now disables the MIME checkbox after checking it if the email is to be sent as HTML
                          2. Fixed a bug in the sample app where the mime type of the root body part would be incorrectly set to an empty string when it should 
                          use the default which is "multipart/mixed". Thanks to Ting L for reporting this bug.
         PJN / 05-01-2014 1. Updated copyright details.
                          2. Updated the code to clean compile on VC 2013
                          3. Fixed a problem with NTLM auth where the code did not correctly set the "m_nSize" parameter value correctly in 
                          CPJNSMPTBase64Encode::Encode & Decode. Thanks to John Pendleton for reporting this bug.
         PJN / 18-01-2013 1. The ConvertToUTF8 method now also handles conversion of ASCII data to UTF8. Thanks to Jean-Christophe Voogden for this update.
                          2. The CPJNSMTPConnection class now provides Set/GetSSLProtocol methods. This allows client code to specify the exact flavour of SSL which
                          the code should speak. Supported protocols are SSL v2/v3, TLS v1.0, TLS v1.1, TLS v1.2 and DTLS v1.0. This is required for some SMTP
                          servers which use the SSLv2 or v3 protocol instead of the more modern TLS v1 protocol. The default is to use TLS v1.0.
                          You may need to use this compatibility setting for the likes of IBM Domino SMTP servers which only support the older SSLv2/v3 setting. 
                          Thanks to Jean-Christophe Voogden for this update.
                          3. Removed all the proxy connection methods as they cannot be easily supported / tested by the author.
                          4. Reworked the code in CPJNSMTPConnection::ConnectESMTP to handle all variants of the 250 response as well as operate case insensitively
                          which is required by the ESMTP RFC.
                          5. Made more methods virtual to facilitate further client customisation
                          6. Fixed an issue in the DoSTARTTLS method where the socket would be left in a connected state if SSL negotiation failed. This would result in
                          later code which sent on the socket during the tear down phase hanging. Thanks to Jean-Christophe Voogden for reporting this issue.
                          7. The sample app is now linked against the latest OpenSSL v1.0.1f dlls.
         PJN / 26-01-2014 1. Updated ASSERT logic at the top of the ConnectESMTP method. Thanks to Jean-Christophe Voogden for reporting this issue.
         PJN / 09-02-2014 1. Fixed a compile problem in the CreateSSLSocket method when the CPJNSMTP_NOSSL preprocessor macro is defined.
                          2. Reworked the logic which does SecureZeroMemory on sensitive string data
         PJN / 13-04-2014 1. The sample app is now linked against the latest OpenSSL v1.0.1g dlls. This version is the patched version of OpenSSL which does not
                          suffer from the Heartbleed bug.
                          2. Please note that by default OpenSSL does not do host name validation. The sample app provided with the PJNSMTP code also does not do
                          host name validation. This means that as it stands the sample app is vulnerable to man in the middle attacks if you use SSL/TLS to connect
                          to a SMTP server. For further information and sample code which you should incorporate into your real SMTP client applications, please see
                          http://wiki.openssl.org/index.php/Hostname_validation, https://github.com/iSECPartners/ssl-conservatory and 
                          http://archives.seul.org/libevent/users/Feb-2013/msg00043.html.
         PJN / 15-11-2014 1. Removed the defunct method CPJNSMTPBodyPart::HexDigit.
                          2. Reworked the CPJNSMTPBodyPart::GetBody method to use ATL::CAtlFile and ATL::CHeapPtr
                          3. CPJNSMTPConnection now takes a static dependency on Wininet.dll instead of using GetProcAddress.
                          4. CPJNSMTPConnection now takes a static dependency on Dnsapi.dll instead of using GetProcAddress.
                          5. The default timeout set in the CPJNSMTPConnection constructor is now 60 seconds for both debug and release builds
                          6. Removed the now defunct CPJNSMTPConnection::MXLookupAvailable method 
                          7. Removed the now defunct CPJNSMTP_NOMXLOOKUP preprocessor value
                          8. Removed the now defunct PJNLoadLibraryFromSystem32.h module from the distribution
                          9. CPJNSMTPConnection::SendMessage has been reworked to use ATL::CAtlFile and ATL::CHeapPtr
                          10. CPJNSMTPConnection::SendMessage now does a UTF-8 conversion on the body of the email when sending
                          a plain email i.e. no HTML or mime if the charset is UTF-8. Thanks to Oliver Pfister for reporting this issue.
                          11. Sample app has been updated to compile cleanly on VS 2013 Update 3 and higher
                          12. The sample app shipped with the source code is now Visual Studio 2008 and as of this release the code is only supported on Visual Studio 
                          2008 and later
                          13. The sample app is now linked against the latest OpenSSL v1.0.1j dlls
         PJN / 18-11-2014 1. Fixed a bug in CPJNSMTPBodyPart::GetBody where file attachments were always being treated as 0 in size. Thanks to Oliver Pfister for 
                          reporting this issue.
         PJN / 14-12-2014 1. Updated the code to use the author's SSLWrappers classes (http://www.naughter.com/sslwrappers.html) to provide the SSL functionality 
                          for PJNSMTP instead of OpenSSL. Please note that the SSLWrappers classes itself depends on the author's CryptoWrappers classes 
                          (http://www.naughter.com/cryptowrappers.html) also. You will need to download both of these libraries and copy their modules into the 
                          PJNSMTP directory. Also note that the SSLWrappers and CryptoWrapper classes are only supported on VC 2013 or later and this means that 
                          the PJNSMTP SSL support is only supported when you compile with VC 2013 or later. The solution files included in the download is now for 
                          VC 2013.
                          2. With the change to using SSLWrappers, the code can now support additional flavours of SSL protocol. They are: SSL v2 on its own, 
                          SSL v3 on its own, The OS SChannel SSL default and AnyTLS. Please see the CPJNSMTPConnection::SSLProtocol enum for more details. The 
                          sample app has been updated to allow these values to be used and the default enum value is now "OSDefault".
         PJN / 16-12-2014 1. Updated the code to use the latest v1.03 version of SSLWrappers
         PJN / 16-01-2015 1. Updated copright details
                          2. CPJNSMTPMessage::m_ReplyTo is now an array instead of a single address. Thanks to Bostjan Erzen for suggesting this update.
         PJN / 25-01-2015 1. Addition of CPJNSMTPBodyPart::InsertChildBodyPart and CPJNSMTPMessage::InsertBodyPart methods. These new methods inserts a body part 
                          as the first element in the hierarchy of child body parts instead of to the end. In addition the CPJNSMTPMessage::AddTextBody 
                          and CPJNSMTPMessage::AddHTMLBody methods now call CPJNSMTPBodyPart::InsertChildBodyPart instead of CPJNSMTPBodyPart::AddChildBodyPart 
                          to ensure that the body text part of the SMTP message appears as the first child body part. This helps avoid issues where Microsoft 
                          Exchange modifies multipart messages where the body text is moved to an attachment called ATT00001: For further information please see 
                          http://kb.mit.edu/confluence/pages/viewpage.action?pageId=4981187 and http://support2.microsoft.com/kb/969854. Thanks to Marco Veldman 
                          for reporting this issue.
                          2. Removed support for DTLSv1 SSL protocol as this is not applicable to SMTP as DTLS is designed for UDP and not TCP. This removal helps
                          resolve an issue where the code did not compile in VC 2013 if you used the "Visual Studio 2013 - Windows XP (v120_xp)" Platform Toolset.
                          Thanks to Bostjan Erzen for reporting this issue.
         PJN / 31-01-2015 1. Fixed an error in CPJNSMTPException::GetErrorMessage where it would not use the full HRESULT value when calling FormatMessage. Thanks
                          to Bostjan Erzen for reporting this issue.
                          2. Improved the error handling in CPJNSMTPConnection::Connect when SSL connection errors occur. Thanks to Bostjan Erzen for reporting 
                          this issue.
                          3. Improved the error handling in CPJNSMTPConnection::DoSTARTTLS when SSL connection errors occur. Thanks to Bostjan Erzen for reporting 
                          this issue.
         PJN / 01-03-2015 1. The DoEHLO method has been updated to check for DSN requirements. This change means that the code will connect with a EHLO instead of 
                          HELO when a DSN is requested. The code will now also check whether DSN's are supported by the server when attempting to send the DSN.
                          If the server does not support DSN's in this scenario then a new custom exception will be thrown by the code in the 
                          CPJNSMTPConnection::FormMailFromCommand method. This ensures the code behaves more consistently with respect to the DSN RFC. Thanks to 
                          Philip Mitchell for prompting this update.
                          2. Change the enum value for CPJNSMTPMessage::DSN_NOT_SPECIFIED to 0 from 0xFFFFFFFF.
         PJN / 29-03-2015 1. Updated CPJNSMTPConnection::AuthLogin to handle more variants of "username:" and "password:" SMTP server responses. Thanks to Wang Le
                          for reporting this issue.
         PJN / 01-06-2015 1. Removed unused CPJNSMTPMessage::ConvertHTMLToPlainText method.
                          2. Removed the linkage in the test app's configuration dialog between sending a message with a HTML body and requiring that the message
                          is MIME encoded. 
                          3. Removed the code from the test app which added a plain text attachment to the message in debug builds.
                          4. Updated CPJNSMTPMessage::AddHTMLBody to handle the case where the message is not MIME encoded. In this case the root body part of the 
                          message is simply set to HTML. Thanks to Thane Hubbell for reporting this issue.
         PJN / 12-05-2016 1. Updated copyright details.
                          2. Minor update to the sample app to refer to "SSL / TLS" in the configuration dialog.
                          3. Verfied the code compiles cleanly using VC 2015.
         PJN / 11-08-2016 1. Reworked the classes to optionally compile without MFC. By default the classes now use STL classes and idioms but if you define 
                          CWSOCKET_MFC_EXTENSTIONS the classes will revert back to the MFC behaviour.
                          2. Reworked the CPJNSMPTBase64Encode class to use ATL::CHeapPtr
                          3. Reworked the CPJNSMPTQPEncode class to use ATL::CHeapPtr
                          4. Reworked the CPJNSMPTQEncode class to use ATL::CHeapPtr
                          5. Added SAL annotations to all the code
                          6. CPJNSMTPBodyPart::ConvertToUTF8 now internally uses ATL::CHeapPtr
                          7. The CPJNSMTPMessage::m_CustomHeaders member variable is now public. As such the 
                          AddCustomHeader, GetCustomHeader, GetNumberOfCustomHeaders & RemoveCustomHeader
                          methods have been removed.
                          8. Removed unnecessary code which catches and rethrows a CPJNSMTPException exception from the version of CPJNSMTPConnection::SendMessage 
                          which sends a message from a memory buffer 
                          9. Reworked the CPJNSMTPMessage::ParseMultipleRecipients method to avoid the use of the C++ new allocator
                          10. Reworked the CPJNSMTPMessage::AddMultipleAttachments method to avoid the use of the C++ new allocator
                          11. Reworked the sample app to exercise the SendMessage method which sends from a memory buffer
         PJN / 17-08-2016 1. Reworked CPJNSMTPConnection class to be inherited from SSLWrappers::CSocket. This allows
                          for easier overloading of the various SSLWrappers class framework virtual functions
         PJN / 19-08-2016 1. Reworked the CPJNSMTPException::GetErrorMessage to not bother putting the last response into the return error message text.
                          2. CPJNSMTPConnection class is now protected inherited from SSLWrappers::CSocket.
         PJN / 15-09-2016 1. Added support for XOAUTH2 authentication. For more information about this SMTP authentication mechanism please see 
                          https://developers.google.com/gmail/xoauth2_protocol.
                          2. Reworked the CPJNSMTPConnection::AuthPlain to provide the userid in the authorization identity value. This is a more compatible form
                          of the auth request and avoids issues using AuthPlain against the gmail SMTP servers.
                          3. Added Initial Client Response (SASL-IR) support to the AuthLogin, AuthPlain & AuthXOAUTH2 methods. For more information about SASL-IR
                          please see https://tools.ietf.org/html/rfc4959.
         PJN / 13-10-2016 1. Fixed a bug in the declaration of the CPJNSMTPConnection class when the CPJNSMTP_NONTLM preprocessor value is defined. Thanks to Chris 
                          Korda for reporting this issue.
                          2. Provided a new overload of the CPJNSMTPBodyPart::AddChildBodyPart method to allow an already heap allocated bodyPart to be added.
                          3. Provided a new overload of the CPJNSMTPBodyPart::InsertChildBodyPart method to allow an already heap allocated bodyPart to be added.
                          4. Made most of the methods of CPJNSMTPBodyPart virtual to allow for easier end user customization of its functionality.
                          5. Provided a new overload of the CPJNSMTPMessage::AddChildBodyPart method to allow an already heap allocated bodyPart to be added.
                          6. Provided a new overload of the CPJNSMTPMessage::InsertChildBodyPart method to allow an already heap allocated bodyPart to be added.
                          7. Made most of the methods of CPJNSMTPMessage virtual to allow for easier end user customization of its functionality.
         PJN / 14-11-2016 1. Removed unnecessary validation of pszUsername & pszPassword parameters in CPJNSMTPConnection::ConnectESMTP. Thanks to Christopher 
                          Craft for reporting this issue.
                          2. Added validation of pszUsername & pszPassword parameters in CPJNSMTPConnection::AuthLogin, CPJNSMTPConnection::AuthPlain &  
                          CPJNSMTPConnection::AuthCramMD5 methods. Thanks to Christopher Craft for reporting this issue.
                          3. Improved the SAL annotations on the CPJNSMTPConnection::ConnectESMTP method.
                          4. Reworked CPJNSMTPBodyPart::ConvertToUTF8 to use ATL::CW2A
                          5. Reworked the non MFC code path of the SMTP code to now be based on a new CPJNSMTP_MFC_EXTENSIONS preprocessor values instead of the 
                          existing CWSOCKET_MFC_EXTENSIONS preprocessor value which not only applies to the socket class
                          6. Optimized the usage of ATL::CT2A & ATL::CA2T throughout the code

Copyright (c) 1998 - 2016 by PJ Naughter (Web: www.naughter.com, Email: pjna@naughter.com)

All rights reserved.

Copyright / Usage Details:

You are allowed to include the source code in any product (commercial, shareware, freeware or otherwise) 
when your product is released in binary form. You are allowed to modify the source code in any way you want 
except you cannot modify the copyright details at the top of each module. If you want to distribute source 
code with your application, then you are only allowed to distribute versions released by the author. This is 
to maintain a single distribution point for the source code. 

Please note that I have been informed that C(PJN)SMTPConnection is being used to develop and send unsolicted bulk mail. 
This was not the intention of the code and the author explicitly forbids use of the code for any software of this kind.

*/


//////////////// Includes /////////////////////////////////////////////////////

#include "stdafx.h"
#include "PJNSmtp.h"
#include "resource.h"
#include "PJNMD5.h"
#ifndef _WININET_
#include <wininet.h>
#pragma message("To avoid this message, please put wininet.h in your pre compiled header (usually stdafx.h)")
#endif //#ifndef _WININET_
#ifndef __ATLENC_H__
#pragma message("To avoid this message, please put atlenc.h in your pre compiled header (usually stdafx.h)")
#include <atlenc.h>
#endif //#ifndef __ATLENC_H__


//////////////// Macros / Locals //////////////////////////////////////////////

#ifdef CPJNSMTP_MFC_EXTENSIONS
#ifdef _DEBUG
#define new DEBUG_NEW
#endif //#ifdef _DEBUG
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

#pragma comment(lib, "wsock32.lib") //Automatically link in the Winsock dll
#pragma comment(lib, "rpcrt4.lib")  //Automatically link in the RPC runtime dll
#pragma comment(lib, "wininet.lib") //Automatically link in the Wininet dll
#pragma comment(lib, "Dnsapi.lib")  //Automatically linke to the DNS dll


//////////////// Implementation ///////////////////////////////////////////////

CPJNSMPTBase64Encode::CPJNSMPTBase64Encode() : m_nSize(0)
{
}

void CPJNSMPTBase64Encode::Encode(_In_reads_bytes_(nSize) const BYTE* pData, _In_ int nSize, _In_ DWORD dwFlags)
{
  //Tidy up any heap memory we have been using
  if (m_Buf.m_pData != NULL)
    m_Buf.Free();

  //Calculate and allocate the buffer to store the encoded data
  m_nSize = ATL::Base64EncodeGetRequiredLength(nSize, dwFlags);
  if (!m_Buf.Allocate(m_nSize + 1)) //We allocate an extra byte so that we can null terminate the result
    CPJNSMTPConnection::ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32);

  //Finally do the encoding
  ATLASSUME(m_Buf.m_pData != NULL);
  if (!ATL::Base64Encode(pData, nSize, m_Buf.m_pData, &m_nSize, dwFlags))
    CPJNSMTPConnection::ThrowPJNSMTPException(IDS_PJNSMTP_FAIL_BASE64_ENCODE, FACILITY_ITF);

  //Null terminate the data
  m_Buf.m_pData[m_nSize] = '\0';
}

void CPJNSMPTBase64Encode::Decode(_In_reads_bytes_(nSize) LPCSTR pData, _In_ int nSize)
{
  //Tidy up any heap memory we have been using
  if (m_Buf.m_pData != NULL)
    m_Buf.Free();

  //Calculate and allocate the buffer to store the encoded data
  m_nSize = ATL::Base64DecodeGetRequiredLength(nSize);
  if (!m_Buf.Allocate(m_nSize + 1)) //We allocate an extra byte so that we can null terminate the result
    CPJNSMTPConnection::ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32);

  //Finally do the encoding
  ATLASSUME(m_Buf.m_pData != NULL);
  if (!ATL::Base64Decode(pData, nSize, reinterpret_cast<BYTE*>(m_Buf.m_pData), &m_nSize))
    CPJNSMTPConnection::ThrowPJNSMTPException(IDS_PJNSMTP_FAIL_BASE64_DECODE, FACILITY_ITF);

  //Null terminate the data
  m_Buf.m_pData[m_nSize] = '\0';
}

void CPJNSMPTBase64Encode::Encode(_In_z_ LPCSTR pszMessage, _In_ DWORD dwFlags)
{
  Encode(reinterpret_cast<const BYTE*>(pszMessage), static_cast<int>(strlen(pszMessage)), dwFlags);
}

void CPJNSMPTBase64Encode::Decode(_In_z_ LPCSTR pszMessage)
{
  Decode(pszMessage, static_cast<int>(strlen(pszMessage)));
}


CPJNSMPTQPEncode::CPJNSMPTQPEncode() : m_nSize(0)
{
}

void CPJNSMPTQPEncode::Encode(_In_reads_bytes_(nSize) const BYTE* pData, _In_ int nSize, _In_ DWORD dwFlags)
{
  //Tidy up any heap memory we have been using
  if (m_Buf.m_pData != NULL)
    m_Buf.Free();

  //Calculate and allocate the buffer to store the encoded data
  m_nSize = ATL::QPEncodeGetRequiredLength(nSize); 
  if (!m_Buf.Allocate(m_nSize + 1)) //We allocate an extra byte so that we can null terminate the result
    CPJNSMTPConnection::ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32);
  
  //Finally do the encoding
  ATLASSUME(m_Buf.m_pData != NULL);
  if (!ATL::QPEncode(const_cast<BYTE*>(pData), nSize, m_Buf.m_pData, &m_nSize, dwFlags)) //ATL::QPEncode is incorrectly defined as taking a BYTE* instead of a const BYTE* for the first parameter
    CPJNSMTPConnection::ThrowPJNSMTPException(IDS_PJNSMTP_FAIL_QP_ENCODE, FACILITY_ITF);

  //Null terminate the data
  m_Buf.m_pData[m_nSize] = '\0';
}

void CPJNSMPTQPEncode::Encode(_In_z_ LPCSTR pszMessage, _In_ DWORD dwFlags)
{
  Encode(reinterpret_cast<const BYTE*>(pszMessage), static_cast<int>(strlen(pszMessage)), dwFlags);
}


CPJNSMPTQEncode::CPJNSMPTQEncode() : m_nSize(0)
{
}

void CPJNSMPTQEncode::Encode(_In_reads_bytes_(nSize) const BYTE* pData, _In_ int nSize, _In_z_ LPCSTR szCharset)
{
  //Tidy up any heap memory we have been using
  if (m_Buf.m_pData != NULL)
    m_Buf.Free();

  //Calculate and allocate the buffer to store the encoded data
  m_nSize = ATL::QEncodeGetRequiredLength(nSize, ATL_MAX_ENC_CHARSET_LENGTH); 
  if (!m_Buf.Allocate(m_nSize + 1)) //We allocate an extra byte so that we can null terminate the result
    CPJNSMTPConnection::ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32);

  //Finally do the encoding
  ATLASSUME(m_Buf.m_pData != NULL);
  if (!ATL::QEncode(const_cast<BYTE*>(pData), nSize, m_Buf.m_pData, &m_nSize, szCharset)) //ATL::QEncode is incorrectly defined as taking a BYTE* instead of a const BYTE* for the first parameter
    CPJNSMTPConnection::ThrowPJNSMTPException(IDS_PJNSMTP_FAIL_Q_ENCODE, FACILITY_ITF);

  //Null terminate the data
  m_Buf.m_pData[m_nSize] = '\0';
}

void CPJNSMPTQEncode::Encode(_In_z_ LPCSTR pszMessage, _In_z_ LPCSTR szCharset)
{
  Encode(reinterpret_cast<const BYTE*>(pszMessage), static_cast<int>(strlen(pszMessage)), szCharset);
}


CPJNSMTPException::CPJNSMTPException(_In_ HRESULT hr, _In_ const CPJNSMTPString& sLastResponse) : m_hr(hr),
                                                                                                  m_sLastResponse(sLastResponse)
{
}

CPJNSMTPException::CPJNSMTPException(_In_ DWORD dwError, _In_ DWORD dwFacility, _In_ const CPJNSMTPString& sLastResponse) : m_hr(MAKE_HRESULT(SEVERITY_ERROR, dwFacility, dwError)),
                                                                                                                            m_sLastResponse(sLastResponse)
{
}

#if _MSC_VER >= 1700
BOOL CPJNSMTPException::GetErrorMessage(_Out_writes_z_(nMaxError) LPTSTR lpszError, _In_ UINT nMaxError, _Out_opt_ PUINT pnHelpContext)
#else	
BOOL CPJNSMTPException::GetErrorMessage(__out_ecount_z(nMaxError) LPTSTR lpszError, __in UINT nMaxError, __out_opt PUINT pnHelpContext)
#endif //#if _MSC_VER >= 1700
{
  //Validate our parameters
  ATLASSERT(lpszError != NULL);
  ATLASSERT(nMaxError != 0);
  lpszError[0] = _T('\0');

  if (pnHelpContext != NULL)
    *pnHelpContext = 0;

  //What will be the return value from this method (assume the worst)
  BOOL bSuccess = FALSE;
    
  if (HRESULT_FACILITY(m_hr) == FACILITY_ITF)
  {
    //Simply load up the string from the string table
    CPJNSMTPString sError;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    bSuccess = sError.LoadString(HRESULT_CODE(m_hr));
    if (bSuccess)
      Checked::tcsncpy_s(lpszError, nMaxError, sError, _TRUNCATE);
  #else
    TCHAR szResource[4096];
    szResource[0] = _T('\0');
    bSuccess = (LoadString(GetModuleHandle(NULL), HRESULT_CODE(m_hr), szResource, sizeof(szResource) / sizeof(TCHAR)) != 0);
    if (bSuccess)
    {
      sError = szResource;
      Checked::tcsncpy_s(lpszError, nMaxError, sError.c_str(), _TRUNCATE);
    }
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
  else
  {
    LPTSTR lpBuffer = NULL;
    DWORD dwReturn = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                                   NULL, m_hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_SYS_DEFAULT),
                                   reinterpret_cast<LPTSTR>(&lpBuffer), 0, NULL);
    if (dwReturn == 0)
      *lpszError = _T('\0');
    else
    {
      bSuccess = TRUE;
      Checked::tcsncpy_s(lpszError, nMaxError, lpBuffer, _TRUNCATE);
      LocalFree(lpBuffer);
    }
  }

  return bSuccess;
}

#ifdef CPJNSMTP_MFC_EXTENSIONS
CPJNSMTPString CPJNSMTPException::GetErrorMessage()
{
  CPJNSMTPString rVal;
  LPTSTR pstrError = rVal.GetBuffer(4096);
  GetErrorMessage(pstrError, 4096, NULL);
  rVal.ReleaseBuffer();
  return rVal;
}

#ifdef _DEBUG
void CPJNSMTPException::Dump(CDumpContext& dc) const
{
  //Let the base class do its thing
  CObject::Dump(dc);

  dc << _T("m_hr = ") << m_hr << _T("\n");
}
#endif //#ifdef _DEBUG

#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

CPJNSMTPAddress::CPJNSMTPAddress() 
{
}

CPJNSMTPAddress::CPJNSMTPAddress(_In_ const CPJNSMTPAddress& address)
{
  *this = address;
}

#ifdef CPJNSMTP_MFC_EXTENSIONS
CPJNSMTPAddress::CPJNSMTPAddress(_In_z_ LPCTSTR pszAddress)
{
  //The local variable which we will operate on
  CPJNSMTPString sTemp(pszAddress);
  sTemp.Trim();

  //divide the substring into friendly names and e-mail addresses
  int nMark = sTemp.Find(_T('<'));
  int nMark2 = sTemp.Find(_T('>'));
  if ((nMark != -1) && (nMark2 != -1) && (nMark2 > (nMark+1)))
  {
    m_sEmailAddress = sTemp.Mid(nMark + 1, nMark2 - nMark - 1);
    m_sFriendlyName = sTemp.Left(nMark);
    
    //Tidy up the parsed values
    m_sFriendlyName.Trim();
    m_sFriendlyName = RemoveQuotes(m_sFriendlyName);
    m_sEmailAddress.Trim();
    m_sEmailAddress = RemoveQuotes(m_sEmailAddress);
  }
  else
  {
    nMark = sTemp.Find(_T('('));
    nMark2 = sTemp.Find(_T(')'));
    if ((nMark != -1) && (nMark2 != -1) && (nMark2 > (nMark+1)))
    {
      m_sEmailAddress = sTemp.Left(nMark);
      m_sFriendlyName = sTemp.Mid(nMark + 1, nMark2 - nMark - 1);
      
      //Tidy up the parsed values
      m_sFriendlyName.Trim();
      m_sFriendlyName = RemoveQuotes(m_sFriendlyName);
      m_sEmailAddress.Trim();
      m_sEmailAddress = RemoveQuotes(m_sEmailAddress);
    }
    else
      m_sEmailAddress = sTemp;
  }
}
#else
CPJNSMTPAddress::CPJNSMTPAddress(_In_z_ LPCTSTR pszAddress)
{
  //The local variable which we will operate on
  CPJNSMTPString sTemp(pszAddress);
  StringTrim(sTemp);

  //divide the substring into friendly names and e-mail addresses
  CPJNSMTPString::size_type nMark = sTemp.find(_T('<'));
  CPJNSMTPString::size_type nMark2 = sTemp.find(_T('>'));
  if ((nMark != CPJNSMTPString::npos) && (nMark2 != CPJNSMTPString::npos) && (nMark2 > (nMark + 1)))
  {
    m_sEmailAddress = sTemp.substr(nMark + 1, nMark2 - nMark - 1); //PJ TO DO: May be an off by one issue here
    m_sFriendlyName = sTemp.substr(0, nMark); //PJ TO DO: May be an off by one issue here

    //Tidy up the parsed values
    StringTrim(m_sFriendlyName);
    m_sFriendlyName = RemoveQuotes(m_sFriendlyName);
    StringTrim(m_sEmailAddress);
    m_sEmailAddress = RemoveQuotes(m_sEmailAddress);
  }
  else
  {
    nMark = sTemp.find(_T('('));
    nMark2 = sTemp.find(_T(')'));
    if ((nMark != CPJNSMTPString::npos) && (nMark2 != CPJNSMTPString ::npos) && (nMark2 > (nMark + 1)))
    {
      m_sEmailAddress = sTemp.substr(0, nMark);
      m_sFriendlyName = sTemp.substr(nMark + 1, nMark2 - nMark - 1);

      //Tidy up the parsed values
      StringTrim(m_sFriendlyName);
      m_sFriendlyName = RemoveQuotes(m_sFriendlyName);
      StringTrim(m_sEmailAddress);
      m_sEmailAddress = RemoveQuotes(m_sEmailAddress);
    }
    else
      m_sEmailAddress = sTemp;
  }
}
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

CPJNSMTPAddress::CPJNSMTPAddress(_In_z_ LPCTSTR pszFriendly, _In_z_ LPCTSTR pszAddress) : m_sFriendlyName(pszFriendly),
                                                                                          m_sEmailAddress(pszAddress) 
{
#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(m_sEmailAddress.GetLength()); //An empty address is not allowed
#else
  ATLASSERT(m_sEmailAddress.length()); //An empty address is not allowed
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

CPJNSMTPAddress& CPJNSMTPAddress::operator=(_In_ const CPJNSMTPAddress& r) 
{ 
  m_sFriendlyName = r.m_sFriendlyName; 
  m_sEmailAddress = r.m_sEmailAddress; 
  return *this;
}

#ifdef CPJNSMTP_MFC_EXTENSIONS
CPJNSMTPAsciiString CPJNSMTPAddress::GetRegularFormat(_In_ BOOL bEncode, _In_ const CPJNSMTPString& sCharset) const
{
  ATLASSERT(m_sEmailAddress.GetLength()); //Email Address must be valid

  CPJNSMTPAsciiString sAddress;
  CPJNSMTPAsciiString sAsciiEmailAddress(m_sEmailAddress);
    
  if (m_sFriendlyName.IsEmpty())
    sAddress = m_sEmailAddress; //Just transfer the address across directly
  else
  {
    if (bEncode)
    {
      CPJNSMTPAsciiString sAsciiEncodedFriendly(CPJNSMTPBodyPart::HeaderEncode(m_sFriendlyName, sCharset));
      sAddress.Format("\"%s\" <%s>", sAsciiEncodedFriendly.operator LPCSTR(), sAsciiEmailAddress.operator LPCSTR());
    }
    else
      sAddress.Format("\"%s\" <%s>", CPJNSMTPAsciiString(m_sFriendlyName).operator LPCSTR(), sAsciiEmailAddress.operator LPCSTR());
  }

  return sAddress;
}

CPJNSMTPString CPJNSMTPAddress::RemoveQuotes(_In_ const CPJNSMTPString& sValue)
{
  //What will be the return value from this method
  CPJNSMTPString sReturn;

  int nLength = sValue.GetLength();
  if (nLength > 2)
  {
    if (sValue.GetAt(0) == _T('"') && sValue.GetAt(nLength - 1) == _T('"'))
      sReturn = sValue.Mid(1, nLength - 2);
    else
      sReturn = sValue;
  }
  else
    sReturn = sValue;

  return sReturn;
}
#else
CPJNSMTPAsciiString CPJNSMTPAddress::GetRegularFormat(_In_ BOOL bEncode, _In_ const CPJNSMTPString& sCharset) const
{
  ATLASSERT(m_sEmailAddress.length()); //Email Address must be valid

  //What will be the return value from this method
  CPJNSMTPAsciiString sAddress;

#ifdef _UNICODE
  ATL::CT2A sAsciiEmailAddress(m_sEmailAddress.c_str());
  LPCSTR pszAsciiEmailAddress = sAsciiEmailAddress;
#else
  LPCSTR pszAsciiEmailAddress = m_sEmailAddress.c_str();
#endif //#ifdef _UNICODE

  if (m_sFriendlyName.length() == 0)
    sAddress = pszAsciiEmailAddress; //Just transfer the address across directly
  else
  {
    std::ostringstream ss;
    if (bEncode)
    {
      CPJNSMTPAsciiString sAsciiEncodedFriendly(CPJNSMTPBodyPart::HeaderEncode(m_sFriendlyName.c_str(), sCharset.c_str()));
      ss << "\"" << sAsciiEncodedFriendly << "\"" << "<" << pszAsciiEmailAddress << ">";
      sAddress = ss.str();
    }
    else
    {
    #ifdef _UNICODE
      ss << "\"" << ATL::CT2A(m_sFriendlyName.c_str()) << "\"" << "<" << pszAsciiEmailAddress << ">";
    #else
      ss << "\"" << m_sFriendlyName.c_str() << "\"" << "<" << pszAsciiEmailAddress << ">";
    #endif //#ifdef _UNICODE
      sAddress = ss.str();
    }
  }

  return sAddress;
}

CPJNSMTPString CPJNSMTPAddress::RemoveQuotes(_In_ const CPJNSMTPString& sValue)
{
  //What will be the return value from this method
  CPJNSMTPString sReturn;

  CPJNSMTPString::size_type nLength = sValue.length();
  if (nLength > 2)
  {
    if (sValue[0] == _T('"') && sValue[nLength - 1] == _T('"'))
      sReturn = sValue.substr(1, nLength - 2);
    else
      sReturn = sValue;
  }
  else
    sReturn = sValue;

  return sReturn;
}
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

#ifndef CPJNSMTP_MFC_EXTENSIONS
void CPJNSMTPAddress::StringTrim(_Inout_ CPJNSMTPString& sString)
{
  CPJNSMTPString::size_type nPos = sString.find_last_not_of(_T(" \n\r\t"));
  if (nPos != CPJNSMTPString::npos)
    sString.resize(nPos + 1);

  nPos = sString.find_first_not_of(_T(" \n\r\t"));
  if (nPos != std::string::npos)
    sString = sString.substr(nPos);
}

void CPJNSMTPAddress::StringReplace(_Inout_ std::string& sString, _In_z_ LPCSTR pszToFind, _In_z_ LPCSTR pszToReplace)
{
  std::string::size_type nFind = sString.find(pszToFind);
  size_t nFindLen = strlen(pszToFind);
  size_t nReplaceLen = strlen(pszToReplace);
  while (nFind != std::string::npos)
  {
    sString.replace(nFind, nFindLen, pszToReplace);
    nFind = sString.find(pszToFind, nFind + nReplaceLen);
  }
}

void CPJNSMTPAddress::StringReplace(_Inout_ std::wstring& sString, _In_z_ LPCWSTR pszToFind, _In_z_ LPCWSTR pszToReplace)
{
  std::wstring::size_type nFind = sString.find(pszToFind);
  size_t nFindLen = wcslen(pszToFind);
  size_t nReplaceLen = wcslen(pszToReplace);
  while (nFind != std::wstring::npos)
  {
    sString.replace(nFind, nFindLen, pszToReplace);
    nFind = sString.find(pszToFind, nFind + nReplaceLen);
  }
}
#endif //#ifndef CPJNSMTP_MFC_EXTENSIONS


CPJNSMTPBodyPart::CPJNSMTPBodyPart() : m_sCharset(_T("iso-8859-1")), 
                                       m_sContentType(_T("text/plain")), 
                                       m_pParentBodyPart(NULL), 
                                       m_bAttachment(FALSE),
                                       m_pszRawBody(NULL),
                                       m_ContentTransferEncoding(NO_ENCODING)
{
  //Automatically generate a unique boundary separator for this body part by creating a guid
  m_sBoundary = CreateGUID();
}

CPJNSMTPBodyPart::CPJNSMTPBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart)
{
  *this = bodyPart;
}

CPJNSMTPBodyPart::~CPJNSMTPBodyPart()
{
  //Free up the array memory
#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<m_ChildBodyParts.GetSize(); i++)
    delete m_ChildBodyParts.GetAt(i);
  m_ChildBodyParts.RemoveAll();
#else
  for (std::vector<CPJNSMTPBodyPart*>::size_type i=0; i<m_ChildBodyParts.size(); i++)
    delete m_ChildBodyParts[i];
  m_ChildBodyParts.clear();
#endif
}

CPJNSMTPBodyPart& CPJNSMTPBodyPart::operator=(_In_ const CPJNSMTPBodyPart& bodyPart)
{
  m_sFilename               = bodyPart.m_sFilename;
  m_sText                   = bodyPart.m_sText;       
  m_sTitle                  = bodyPart.m_sTitle;      
  m_sContentType            = bodyPart.m_sContentType;
  m_sCharset                = bodyPart.m_sCharset;
  m_sContentID              = bodyPart.m_sContentID;
  m_sContentLocation        = bodyPart.m_sContentLocation;
  m_pParentBodyPart         = bodyPart.m_pParentBodyPart;
  m_sBoundary               = bodyPart.m_sBoundary;
  m_bAttachment             = bodyPart.m_bAttachment;
  m_pszRawBody              = bodyPart.m_pszRawBody;
  m_ContentTransferEncoding = bodyPart.m_ContentTransferEncoding;

  //Free up the array memory
#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<m_ChildBodyParts.GetSize(); i++)
    delete m_ChildBodyParts.GetAt(i);
  m_ChildBodyParts.RemoveAll();

  //Now copy over the new object
  for (INT_PTR i=0; i<bodyPart.m_ChildBodyParts.GetSize(); i++)
  {
    CPJNSMTPBodyPart* pBodyPart = new CPJNSMTPBodyPart(*bodyPart.m_ChildBodyParts[i]);
    pBodyPart->m_pParentBodyPart = this;
    m_ChildBodyParts.Add(pBodyPart);
  }
#else
  for (std::vector<CPJNSMTPBodyPart*>::size_type i=0; i<m_ChildBodyParts.size(); i++)
    delete m_ChildBodyParts[i];
  m_ChildBodyParts.clear();

  //Now copy over the new object
  for (std::vector<CPJNSMTPBodyPart*>::size_type i=0; i<bodyPart.m_ChildBodyParts.size(); i++)
  {
    CPJNSMTPBodyPart* pBodyPart = new CPJNSMTPBodyPart(*bodyPart.m_ChildBodyParts[i]);
    pBodyPart->m_pParentBodyPart = this;
    m_ChildBodyParts.push_back(pBodyPart);
  }
#endif

  return *this;
}

CPJNSMTPString CPJNSMTPBodyPart::CreateGUID()
{
  UUID uuid;
  memset(&uuid, 0, sizeof(uuid));
  RPC_STATUS status = UuidCreate(&uuid);
  if ((status != RPC_S_OK) && (status != RPC_S_UUID_LOCAL_ONLY))
    CPJNSMTPConnection::ThrowPJNSMTPException(status, FACILITY_RPC);
  
  //Convert it to a string
#ifdef _UNICODE
  RPC_WSTR pszGuid = NULL;
#else
  RPC_CSTR pszGuid = NULL;
#endif //#ifdef _UNICODE
#pragma warning(suppress: 6102)
  status = UuidToString(&uuid, &pszGuid);
  if (status != RPC_S_OK)
    CPJNSMTPConnection::ThrowPJNSMTPException(status, FACILITY_RPC);

  //What will be the return value from this method
  CPJNSMTPString sGUID(reinterpret_cast<TCHAR*>(pszGuid));

  //Free up the temp memory
  RpcStringFree(&pszGuid);

  return sGUID;
}

BOOL CPJNSMTPBodyPart::SetFilename(_In_ const CPJNSMTPString& sFilename)
{
  //Validate our parameters
#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(sFilename.GetLength());  //You need to a valid Filename!
#else
  ATLASSERT(sFilename.length());  //You need to a valid Filename!
#endif //#ifndef CPJNSMTP_MFC_EXTENSIONS

  //Hive away the filename and form the title from the filename
  TCHAR sPath[_MAX_PATH];
  TCHAR sFname[_MAX_FNAME];
  TCHAR sExt[_MAX_EXT];
#ifdef CPJNSMTP_MFC_EXTENSIONS
  _tsplitpath_s(sFilename, NULL, 0, NULL, 0, sFname, sizeof(sFname)/sizeof(TCHAR), sExt, sizeof(sExt)/sizeof(TCHAR));
#else
  _tsplitpath_s(sFilename.c_str(), NULL, 0, NULL, 0, sFname, sizeof(sFname) / sizeof(TCHAR), sExt, sizeof(sExt) / sizeof(TCHAR));
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  _tmakepath_s(sPath, sizeof(sPath)/sizeof(TCHAR), NULL, NULL, sFname, sExt);
  m_sFilename = sFilename;
  m_sTitle = sPath;

  //Also set the content type to be appropiate for an attachment
  m_sContentType = _T("application/octet-stream");

  //Also set the Content-Transfer-Encoding to base64
  SetContentTransferEncoding(BASE64_ENCODING);
  
  //Also set the attachment setting
  m_bAttachment = TRUE;

  return TRUE;
}

void CPJNSMTPBodyPart::SetRawBody(_In_z_ LPCSTR pszRawBody)
{
  m_pszRawBody = pszRawBody;
}

LPCSTR CPJNSMTPBodyPart::GetRawBody()
{
  return m_pszRawBody;
}

void CPJNSMTPBodyPart::SetText(_In_ const CPJNSMTPString& sText)
{
  m_sText = sText;

  //Ensure lines are correctly wrapped
#ifdef CPJNSMTP_MFC_EXTENSIONS
  m_sText.Replace(_T("\r\n"), _T("\n"));
  m_sText.Replace(_T("\r"), _T("\n"));
  m_sText.Replace(_T("\n"), _T("\r\n"));
#else
  CPJNSMTPAddress::StringReplace(m_sText, _T("\r\n"), _T("\n"));
  CPJNSMTPAddress::StringReplace(m_sText, _T("\r"), _T("\n"));
  CPJNSMTPAddress::StringReplace(m_sText, _T("\n"), _T("\r\n"));
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //Also set the content type while we are at it
  m_sContentType = _T("text/plain");
}

void CPJNSMTPBodyPart::SetContentID(_In_z_ LPCTSTR pszContentID)
{
#ifdef CPJNSMTP_MFC_EXTENSIONS
  m_sContentLocation.Empty();
#else
  m_sContentLocation.clear();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  m_sContentID = pszContentID; 
}

CPJNSMTPString CPJNSMTPBodyPart::GetContentID() const
{ 
  return m_sContentID; 
}

void CPJNSMTPBodyPart::SetContentLocation(_In_z_ LPCTSTR pszContentLocation)
{ 
#ifdef CPJNSMTP_MFC_EXTENSIONS
  m_sContentID.Empty();
#else
  m_sContentID.clear();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  m_sContentLocation = pszContentLocation; 
}

CPJNSMTPString CPJNSMTPBodyPart::GetContentLocation() const
{ 
  return m_sContentLocation; 
}

void CPJNSMTPBodyPart::SetAttachment(_In_ BOOL bAttachment)
{ 
  m_bAttachment = bAttachment; 
  if (m_bAttachment)
    SetContentTransferEncoding(BASE64_ENCODING);
}

CPJNSMTPAsciiString CPJNSMTPBodyPart::ConvertToUTF8(_In_z_ LPCTSTR pszText)
{
  //What will be the return value from this method
#ifdef _UNICODE
  return CPJNSMTPAsciiString(ATL::CW2A(pszText, CP_UTF8));
#else
  return CPJNSMTPAsciiString(ATL::CW2A(ATL::CA2W(pszText), CP_UTF8));
#endif
}

CPJNSMTPAsciiString CPJNSMTPBodyPart::GetHeader(_In_ const CPJNSMTPString& sRootCharset)
{
  //Validate our parameters
  ATLASSERT(m_pParentBodyPart != NULL);

  CPJNSMTPString sHeaderT;
  if (m_bAttachment)
  {
    //Ok, it's an attachment
    
    //Check if filename title has non-ASCII characters
    BOOL bNeedTitleEscaping = FALSE;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    for (int i=0; i<m_sTitle.GetLength() && !bNeedTitleEscaping; i++)
  #else
    for (CPJNSMTPString::size_type i=0; i<m_sTitle.length() && !bNeedTitleEscaping; i++)
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    {
      TCHAR cChar = m_sTitle[i];
      if (cChar <  _T(' ') || cChar > _T('~')) 
        bNeedTitleEscaping = TRUE;
    }
    //Encode the title of the file if required
    CPJNSMTPString sTitle;
    if (bNeedTitleEscaping)
      sTitle = HeaderEncodeT(m_sTitle, sRootCharset);
    else
      sTitle = m_sTitle;

    //Form the header to go along with this body part
    if (GetNumberOfChildBodyParts())
    {
      switch (m_ContentTransferEncoding)
      {
        case BASE64_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n\r\n--%s\r\nContent-Type: %s; name=\"%s\"; Boundary=\"%s\"\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename=\"%s\"\r\n"), 
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), sTitle.operator LPCTSTR(), m_sBoundary.operator LPCTSTR(), sTitle.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; name=\"") << sTitle << _T("\"; Boundary=\"") << m_sBoundary << _T("\"\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename=\"") << sTitle << _T("\"\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        case QP_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n\r\n--%s\r\nContent-Type: %s; name=\"%s\"; Boundary=\"%s\"\r\nContent-Transfer-Encoding: quoted-printable\r\nContent-Disposition: attachment; filename=\"%s\"\r\n"), 
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), sTitle.operator LPCTSTR(), m_sBoundary.operator LPCTSTR(), sTitle.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; name=\"") << sTitle << _T("\"; Boundary=\"") << m_sBoundary << _T("\"\r\nContent-Transfer-Encoding: quoted-printable\r\nContent-Disposition: attachment; filename=\"") << sTitle << _T("\"\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        default:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n\r\n--%s\r\nContent-Type: %s; name=\"%s\"; Boundary=\"%s\"\r\nContent-Disposition: attachment; filename=\"%s\"\r\n"), 
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), sTitle.operator LPCTSTR(), m_sBoundary.operator LPCTSTR(), sTitle.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; name=\"") << sTitle << _T("\"; Boundary=\"") << m_sBoundary << _T("\"\r\nContent-Disposition: attachment; filename=\"") << sTitle << _T("\"\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
      }
    }
    else
    {
      switch (m_ContentTransferEncoding)
      {
        case BASE64_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n\r\n--%s\r\nContent-Type: %s; name=\"%s\"\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename=\"%s\"\r\n"), 
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), sTitle.operator LPCTSTR(), sTitle.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; name=\"") << sTitle << _T("\"\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename=\"") << sTitle << _T("\"\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        case QP_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n\r\n--%s\r\nContent-Type: %s; name=\"%s\"\r\nContent-Transfer-Encoding: quoted-printable\r\nContent-Disposition: attachment; filename=\"%s\"\r\n"), 
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), sTitle.operator LPCTSTR(), sTitle.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; name=\"") << sTitle << _T("\"\r\nContent-Transfer-Encoding: quoted-printable\r\nContent-Disposition: attachment; filename=\"") << sTitle << _T("\"\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        default:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n\r\n--%s\r\nContent-Type: %s; name=\"%s\"\r\nContent-Disposition: attachment; filename=\"%s\"\r\n"), 
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), sTitle.operator LPCTSTR(), sTitle.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; name=\"") << sTitle << _T("\"\r\nContent-Disposition: attachment; filename=\"") << sTitle << _T("\"\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
      }
    }
  }
  else
  {
    //ok, it's not an attachment

    //Form the header to go along with this body part
    if (GetNumberOfChildBodyParts())
    {
      switch (m_ContentTransferEncoding)
      {
        case BASE64_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n--%s\r\nContent-Type: %s; charset=%s; Boundary=\"%s\"\r\nContent-Transfer-Encoding: base64\r\n"),
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), m_sCharset.operator LPCTSTR(), m_sBoundary.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; charset=") << m_sCharset << _T("; Boundary=\"") << m_sBoundary << _T("\"\r\nContent-Transfer-Encoding: base64\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        case QP_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n--%s\r\nContent-Type: %s; charset=%s; Boundary=\"%s\"\r\nContent-Transfer-Encoding: quoted-printable\r\n"),
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), m_sCharset.operator LPCTSTR(), m_sBoundary.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; charset=") << m_sCharset << _T("; Boundary=\"") << m_sBoundary << _T("\"\r\nContent-Transfer-Encoding: quoted-printable\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        default:
        {  
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n--%s\r\nContent-Type: %s; charset=%s; Boundary=\"%s\"\r\n"),
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), m_sCharset.operator LPCTSTR(), m_sBoundary.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; charset=") << m_sCharset << _T("; Boundary=\"") << m_sBoundary << _T("\"\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
      }
    }
    else
    {
      switch (m_ContentTransferEncoding)
      {
        case BASE64_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n--%s\r\nContent-Type: %s; charset=%s\r\nContent-Transfer-Encoding: base64\r\n"),
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), m_sCharset.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; charset=") << m_sCharset << _T("\r\nContent-Transfer-Encoding: base64\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        case QP_ENCODING:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n--%s\r\nContent-Type: %s; charset=%s\r\nContent-Transfer-Encoding: quoted-printable\r\n"),
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), m_sCharset.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; charset=") << m_sCharset << _T("\r\nContent-Transfer-Encoding: quoted-printable\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
        default:
        {
        #ifdef CPJNSMTP_MFC_EXTENSIONS
          sHeaderT.Format(_T("\r\n--%s\r\nContent-Type: %s; charset=%s\r\n"),
                          m_pParentBodyPart->m_sBoundary.operator LPCTSTR(), m_sContentType.operator LPCTSTR(), m_sCharset.operator LPCTSTR());
        #else
        #ifdef _UNICODE
          std::wostringstream ss;
        #else
          std::ostringstream ss;
        #endif //#ifdef _UNICODE
          ss << _T("\r\n--") << m_pParentBodyPart->m_sBoundary << _T("\r\nContent-Type: ") << m_sContentType << _T("; charset=") << m_sCharset << _T("\r\n");
          sHeaderT = ss.str();
        #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
          break;
        }
      }
    }
  }

  //Add the other headers
#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (m_sContentID.GetLength())
#else
  if (m_sContentID.length())
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  {
    CPJNSMTPString sLine(_T("Content-ID: "));
    sLine += m_sContentID;
    sLine += _T("\r\n");
    sHeaderT += sLine;
  }
#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (m_sContentLocation.GetLength())
#else
  if (m_sContentLocation.length())
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  {
    CPJNSMTPString sLine(_T("Content-Location: "));
    sLine += m_sContentLocation;
    sLine += _T("\r\n");
    sHeaderT += sLine;
  }
  sHeaderT += _T("\r\n");

#ifdef CPJNSMTP_MFC_EXTENSIONS
  return CPJNSMTPAsciiString(sHeaderT);
#else
#ifdef _UNICODE
  return CPJNSMTPAsciiString(ATL::CT2A(sHeaderT.c_str()));
#else
  return CPJNSMTPAsciiString(sHeaderT);
#endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

CPJNSMTPAsciiString CPJNSMTPBodyPart::GetBody(_In_ BOOL bDoSingleDotFix)
{
  //if the body is text we must convert it to the declared encoding, this could create some
  //problems since Windows conversion functions (such as WideCharToMultiByte) uses UINT
  //code page and not a String like HTTP uses.
  //For now we will add the support for UTF-8, to support other encodings we have to 
  //implement a mapping between the "internet name" of the encoding and its LCID(UINT)

  //What will be the return value from this method
  CPJNSMTPAsciiString sBodyA;
  
  if (m_pszRawBody != NULL)
    sBodyA = m_pszRawBody;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  else if (m_sFilename.GetLength())
#else
  else if (m_sFilename.length())
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  {
    //Ok, it's a file  
    ATL::CAtlFile file;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    HRESULT hr = file.Create(m_sFilename, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
  #else
    HRESULT hr = file.Create(m_sFilename.c_str(), GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    if (FAILED(hr))
      CPJNSMTPConnection::ThrowPJNSMTPException(hr);
    
    //Get the length of the file (we only support sending files less than 2GB!)
    ULONGLONG nFileSize = 0;
    hr = file.GetSize(nFileSize);
    if (FAILED(hr))
      CPJNSMTPConnection::ThrowPJNSMTPException(hr);
    if (nFileSize >= INT_MAX)
      CPJNSMTPConnection::ThrowPJNSMTPException(IDS_PJNSMTP_FILE_SIZE_TO_SEND_TOO_LARGE, FACILITY_ITF);
    
    //Allocate some memory for the contents of the file
    ATL::CHeapPtr<BYTE> body;
    if (!body.Allocate(static_cast<size_t>(nFileSize + 1)))
      CPJNSMTPConnection::ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32); 

    //Read in the contents of the file
    DWORD dwBytesWritten = 0;
    hr = file.Read(body, static_cast<DWORD>(nFileSize), dwBytesWritten);
    if (FAILED(hr))
      CPJNSMTPConnection::ThrowPJNSMTPException(hr);

    switch (m_ContentTransferEncoding)
    {
      case BASE64_ENCODING:
      {
        //Do the encoding
        CPJNSMPTBase64Encode encode;
        encode.Encode(body, dwBytesWritten, ATL_BASE64_FLAG_NONE);

        //Form the body for this body part
        sBodyA = encode.Result();
        break;
      }
      case QP_ENCODING:
      {
        //Do the encoding
        CPJNSMPTQPEncode encode;
        encode.Encode(body, dwBytesWritten, 0);

        //Form the body for this body part
        sBodyA = encode.Result();
        break;
      }
      default:
      {
        //NULL terminate the data before we try to treat it as a char*
        body.m_pData[nFileSize] = '\0';

        //The body of this body part is just the raw file contents
        sBodyA = reinterpret_cast<const char*>(body.m_pData);
        break;
      }
    }
  }
  else
  {
    switch (m_ContentTransferEncoding)
    {
      case BASE64_ENCODING:
      {
        //Do the UTF8 conversion if necessary
        CPJNSMTPAsciiString sBuff;
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        if (m_sCharset.CompareNoCase(_T("UTF-8")) == 0)
          sBuff = ConvertToUTF8(m_sText);
        else
          sBuff = m_sText;
      #else   
        if (_tcsicmp(m_sCharset.c_str(), _T("UTF-8")) == 0)
          sBuff = ConvertToUTF8(m_sText.c_str());
        else
        {
        #ifdef _UNICODE
          sBuff = ATL::CT2A(m_sText.c_str());
        #else
          sBuff = m_sText;
        #endif //#ifdef _UNICODE
        }
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
        
        //Do the encoding
        CPJNSMPTBase64Encode encode;
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        encode.Encode(sBuff, ATL_BASE64_FLAG_NONE);
      #else
        encode.Encode(sBuff.c_str(), ATL_BASE64_FLAG_NONE);
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

        //Form the body for this body part
        sBodyA = encode.Result();
        break;
      }
      case QP_ENCODING:
      {
        //Do the UTF8 conversion if necessary
        CPJNSMTPAsciiString sBuff;
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        if (m_sCharset.CompareNoCase(_T("UTF-8")) == 0)
          sBuff = ConvertToUTF8(m_sText);
        else
          sBuff = m_sText;
      #else   
        if (_tcsicmp(m_sCharset.c_str(), _T("UTF-8")) == 0)
          sBuff = ConvertToUTF8(m_sText.c_str());
        else
        {
        #ifdef _UNICODE
          sBuff = ATL::CT2A(m_sText.c_str());
        #else
          sBuff = m_sText;
        #endif //#ifdef _UNICODE
        }
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

        //Do the encoding
        CPJNSMPTQPEncode encode;
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        encode.Encode(sBuff, 0);
      #else
        encode.Encode(sBuff.c_str(), 0);
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
        sBodyA = encode.Result();
        if (bDoSingleDotFix)
          FixSingleDot(sBodyA);
        break;
      }
      default:
      {
        //Do the UTF8 conversion if necessary
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        if (m_sCharset.CompareNoCase(_T("UTF-8")) == 0)
          sBodyA = ConvertToUTF8(m_sText);
        else
          sBodyA = m_sText;
      #else   
        if (_tcsicmp(m_sCharset.c_str(), _T("UTF-8")) == 0)
          sBodyA = ConvertToUTF8(m_sText.c_str());
        else
        {
        #ifdef _UNICODE
          sBodyA = ATL::CT2A(m_sText.c_str());
        #else
          sBodyA = m_sText;
        #endif //#ifdef _UNICODE
        }
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

        //No encoding to do
        if (bDoSingleDotFix)
          FixSingleDot(sBodyA);
        break;
      }
    }
  }

  return sBodyA;
}

CPJNSMTPAsciiString CPJNSMTPBodyPart::GetFooter()
{
  //Form the MIME footer
  CPJNSMTPAsciiString sFooter;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  sFooter.Format("\r\n--%s--", CPJNSMTPAsciiString(m_sBoundary).operator LPCSTR());
#else
  std::ostringstream ss;
#ifdef _UNICODE
  ss << "\r\n--" << ATL::CT2A(m_sBoundary.c_str()) << "--";
#else
  ss << "\r\n--" << m_sBoundary << "--";
#endif //#ifdef _UNICODE
  sFooter = ss.str();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  return sFooter;
}

INT_PTR CPJNSMTPBodyPart::GetNumberOfChildBodyParts() const
{
#ifdef CPJNSMTP_MFC_EXTENSIONS
  return m_ChildBodyParts.GetSize();
#else
  return m_ChildBodyParts.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

INT_PTR CPJNSMTPBodyPart::AddChildBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart)
{
  CPJNSMTPBodyPart* pNewBodyPart = new CPJNSMTPBodyPart(bodyPart);
  return AddChildBodyPart(pNewBodyPart);
}

INT_PTR CPJNSMTPBodyPart::AddChildBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart)
{
  pBodyPart->m_pParentBodyPart = this;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(m_sContentType.GetLength()); //Did you forget to call SetContentType
  return m_ChildBodyParts.Add(pBodyPart);
#else
  ATLASSERT(m_sContentType.length()); //Did you forget to call SetContentType
  m_ChildBodyParts.push_back(pBodyPart);
  return m_ChildBodyParts.size() - 1;
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

void CPJNSMTPBodyPart::RemoveChildBodyPart(_In_ INT_PTR nIndex)
{
  CPJNSMTPBodyPart* pBodyPart = m_ChildBodyParts[nIndex];
  delete pBodyPart;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  m_ChildBodyParts.RemoveAt(nIndex);
#else
  m_ChildBodyParts.erase(m_ChildBodyParts.begin() + nIndex);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

void CPJNSMTPBodyPart::InsertChildBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart)
{
  CPJNSMTPBodyPart* pNewBodyPart = new CPJNSMTPBodyPart(bodyPart);
  InsertChildBodyPart(pNewBodyPart);
}

void CPJNSMTPBodyPart::InsertChildBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart)
{
  pBodyPart->m_pParentBodyPart = this;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(m_sContentType.GetLength()); //Did you forget to call SetContentType
  m_ChildBodyParts.InsertAt(0, pBodyPart);
#else
  ATLASSERT(m_sContentType.length()); //Did you forget to call SetContentType
  m_ChildBodyParts.insert(m_ChildBodyParts.begin(), pBodyPart);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

CPJNSMTPBodyPart* CPJNSMTPBodyPart::GetChildBodyPart(_In_ INT_PTR nIndex)
{
  return m_ChildBodyParts[nIndex];
}

CPJNSMTPBodyPart* CPJNSMTPBodyPart::GetParentBodyPart()
{
  return m_pParentBodyPart;
}

void CPJNSMTPBodyPart::FixSingleDot(_Inout_ CPJNSMTPAsciiString& sBody)
{
  //Handle a dot on the first line on its own (by replacing it with a double dot as referred to in 
  //section 4.5.2 of the SMTP RFC)  
#ifdef CPJNSMTP_MFC_EXTENSIONS
  if ((sBody.Find(".\r\n") == 0) || (sBody == "."))
    sBody.Insert(0, ".");

  //Also handle any subsequent single dot lines
  sBody.Replace("\r\n.", "\r\n..");
#else
  if ((sBody.find(".\r\n") == 0) || (sBody == "."))
    sBody.insert(0, ".");

  //Also handle any subsequent single dot lines
  CPJNSMTPAddress::StringReplace(sBody, "\r\n.", "\r\n..");
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

CPJNSMTPBodyPart* CPJNSMTPBodyPart::FindFirstBodyPart(_In_ const CPJNSMTPString& sContentType)
{
#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<m_ChildBodyParts.GetSize(); i++)
#else
  for (std::vector<CPJNSMTPBodyPart*>::size_type i=0; i<m_ChildBodyParts.size(); i++)
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  {
    CPJNSMTPBodyPart* pBodyPart = m_ChildBodyParts[i];
    if (pBodyPart->m_sContentType == sContentType)
      return pBodyPart;
  }
  return NULL;
}

CPJNSMTPAsciiString CPJNSMTPBodyPart::HeaderEncode(_In_ const CPJNSMTPString& sText, _In_ const CPJNSMTPString& sCharset)
{
  CPJNSMPTQEncode encode;

  //Do the UTF8 conversion if necessary
  CPJNSMTPAsciiString sLocalText;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (sCharset.CompareNoCase(_T("UTF-8")) == 0)
    sLocalText = CPJNSMTPBodyPart::ConvertToUTF8(sText);
  else
    sLocalText = sText;
  encode.Encode(sLocalText, CPJNSMTPAsciiString(sCharset));
#else
  if (_tcsicmp(sCharset.c_str(), _T("UTF-8")) == 0)
    sLocalText = CPJNSMTPBodyPart::ConvertToUTF8(sText.c_str());
  else
  {
  #ifdef _UNICODE
    sLocalText = ATL::CT2A(sText.c_str());
  #else
    sLocalText = sText;
  #endif //#ifdef _UNICODE
  }
#ifdef _UNICODE
  encode.Encode(sLocalText.c_str(), ATL::CT2A(sCharset.c_str()));
#else
  encode.Encode(sLocalText.c_str(), sCharset.c_str());
#endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  return CPJNSMTPAsciiString(encode.Result());
}

CPJNSMTPString CPJNSMTPBodyPart::HeaderEncodeT(_In_ const CPJNSMTPString& sText, _In_ const CPJNSMTPString& sCharset)
{
#ifdef CPJNSMTP_MFC_EXTENSIONS
  return CPJNSMTPString(HeaderEncode(sText, sCharset));
#else
#ifdef _UNICODE
  return CPJNSMTPString(ATL::CA2T(HeaderEncode(sText, sCharset).c_str()));
#else
  return CPJNSMTPString(HeaderEncode(sText, sCharset));
#endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

//Note that this method does not absolutely ensure each line is a maximum of 78 characters
//long because if each line needs to be Q encoded, this will result in a variable expansion size 
//of the resultant encoded line. The function will most definitely not exceed the 998 
//characters limit though
CPJNSMTPAsciiString CPJNSMTPBodyPart::FoldSubjectHeader(_In_ const CPJNSMTPString& sSubject, _In_ const CPJNSMTPString& sCharset)
{
  //Do the UTF8 conversion if necessary
  CPJNSMTPAsciiString sLocalSubject;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (sCharset.CompareNoCase(_T("UTF-8")) == 0)
    sLocalSubject = CPJNSMTPBodyPart::ConvertToUTF8(sSubject);
  else
    sLocalSubject = sSubject;
#else
  if (_tcsicmp(sCharset.c_str(), _T("UTF-8")) == 0)
    sLocalSubject = CPJNSMTPBodyPart::ConvertToUTF8(sSubject.c_str());
  else
  {
  #ifdef _UNICODE
    sLocalSubject = ATL::CT2A(sSubject.c_str());
  #else
    sLocalSubject = sSubject;
  #endif //#ifdef _UNICODE
  }
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //What will be the return value from this method
  CPJNSMTPAsciiString sOutputHeader("Subject: ");

  //Create an ASCII version of the charset string if required in the loop
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiCharset(sCharset);
  int nCharsetLength = sAsciiCharset.GetLength();
#else

#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiCharset(ATL::CT2A(sCharset.c_str()));
#else
  CPJNSMTPAsciiString sAsciiCharset(sCharset);
#endif //#ifdef _UNICODE
  int nCharsetLength = static_cast<int>(sAsciiCharset.length());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //Fold the input string to a maximum of 78 characters per line
#ifdef CPJNSMTP_MFC_EXTENSIONS
  LPCSTR pszLine = sLocalSubject.operator LPCSTR();
  int nHeaderLeftLength = sLocalSubject.GetLength();
#else
  LPCSTR pszLine = sLocalSubject.c_str();
  int nHeaderLeftLength = static_cast<int>(sLocalSubject.length());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  int nLineNumber = 0;
  while (nHeaderLeftLength > 78)
  {
    //Remember what line number we are working on. This is required 
    ++nLineNumber;
  
    //Calculate the line length to use for this line
    int nLineLength = 0;
    int nMaxLineLength = 78 - nCharsetLength - 7 - 9;
    if (nLineNumber != 1)
      --nMaxLineLength;
    ATLASSERT(nMaxLineLength > 0);

    if (nHeaderLeftLength > nMaxLineLength)
      nLineLength = nMaxLineLength;
    else
      nLineLength = nHeaderLeftLength;
    
    //copy the current line to a local buffer
    char szLineBuf[79];
    ATLASSUME(nLineLength < 79);
    strncpy_s(szLineBuf, sizeof(szLineBuf), pszLine, nLineLength);
    szLineBuf[nLineLength] = '\0';

    //Prepare for the next time around
    pszLine += nLineLength;
    nHeaderLeftLength -= nLineLength;

    //Add the current line to the output parameter
    CPJNSMPTQEncode encode;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    encode.Encode(szLineBuf, sAsciiCharset);
  #else
    encode.Encode(szLineBuf, sAsciiCharset.c_str());
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    sOutputHeader += encode.Result();
    
    //And prepare for folding the next line if we still have some header to process
    if (nHeaderLeftLength)
      sOutputHeader += "\r\n ";  
  }
  
  //Finish up the remaining text left
  if (nHeaderLeftLength)
  {
    //Add the last line
    CPJNSMPTQEncode encode;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    encode.Encode(pszLine, sAsciiCharset);
  #else
    encode.Encode(pszLine, sAsciiCharset.c_str());
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    sOutputHeader += encode.Result();
  }
  
  return sOutputHeader;
}


CPJNSMTPMessage::CPJNSMTPMessage() : m_sXMailer(_T("CPJNSMTPConnection v3.23")), 
                                     m_bMime(FALSE), 
                                     m_Priority(NoPriority),
                                     m_DSNReturnType(HeadersOnly),
                                     m_dwDSN(DSN_NOT_SPECIFIED),
                                     m_bAddressHeaderEncoding(FALSE)
{
}

CPJNSMTPMessage::CPJNSMTPMessage(_In_ const CPJNSMTPMessage& message)
{
  *this = message;
}

CPJNSMTPMessage& CPJNSMTPMessage::operator=(_In_ const CPJNSMTPMessage& message) 
{ 
  m_From                   = message.m_From;
  m_sSubject               = message.m_sSubject;
  m_sXMailer               = message.m_sXMailer;
  m_RootPart               = message.m_RootPart;
  m_Priority               = message.m_Priority;
  m_DSNReturnType          = message.m_DSNReturnType;
  m_dwDSN                  = message.m_dwDSN;
  m_sENVID                 = message.m_sENVID;
  m_bMime                  = message.m_bMime;
  m_bAddressHeaderEncoding = message.m_bAddressHeaderEncoding;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  m_To.Copy(message.m_To);
  m_CC.Copy(message.m_CC);
  m_BCC.Copy(message.m_BCC);
  m_ReplyTo.Copy(message.m_ReplyTo);
  m_MessageDispositionEmailAddresses.Copy(message.m_MessageDispositionEmailAddresses);
  m_CustomHeaders.Copy(message.m_CustomHeaders);
#else
  m_To                               = message.m_To;
  m_CC                               = message.m_CC;
  m_BCC                              = message.m_BCC;
  m_ReplyTo                          = message.m_ReplyTo;
  m_MessageDispositionEmailAddresses = message.m_MessageDispositionEmailAddresses;
  m_CustomHeaders                    = message.m_CustomHeaders;
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  return *this;
}

void CPJNSMTPMessage::SetCharset(_In_z_ LPCTSTR sCharset)
{
  m_RootPart.SetCharset(sCharset);
}

CPJNSMTPString CPJNSMTPMessage::GetCharset() const
{
  return m_RootPart.GetCharset();
}

INT_PTR CPJNSMTPMessage::AddBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart)
{
  SetMime(TRUE); //Body parts implies Mime
  return m_RootPart.AddChildBodyPart(bodyPart);
}

INT_PTR CPJNSMTPMessage::AddBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart)
{
  SetMime(TRUE); //Body parts implies Mime
  return m_RootPart.AddChildBodyPart(pBodyPart);
}

void CPJNSMTPMessage::InsertBodyPart(_In_ const CPJNSMTPBodyPart& bodyPart)
{
  SetMime(TRUE); //Body parts implies Mime
  return m_RootPart.InsertChildBodyPart(bodyPart);
}

void CPJNSMTPMessage::InsertBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart)
{
  SetMime(TRUE); //Body parts implies Mime
  return m_RootPart.InsertChildBodyPart(pBodyPart);
}

void CPJNSMTPMessage::RemoveBodyPart(_In_ INT_PTR nIndex)
{
  m_RootPart.RemoveChildBodyPart(nIndex);
}

CPJNSMTPBodyPart* CPJNSMTPMessage::GetBodyPart(_In_ INT_PTR nIndex)
{
  return m_RootPart.GetChildBodyPart(nIndex);
}

INT_PTR CPJNSMTPMessage::GetNumberOfBodyParts() const
{
  return m_RootPart.GetNumberOfChildBodyParts();
}

CPJNSMTPAsciiString CPJNSMTPMessage::FormDateHeader()
{
  //The static lookup arrays we use to form the day of week and month strings
  static const char* pszDOW[] = 
  {
    "Sun",
    "Mon",
    "Tue",
    "Wed",
    "Thu",
    "Fri",
    "Sat"
  };
  
  static const char* pszMonth[] = 
  {
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"
  };
  
  //Form the Timezone info which will form part of the Date header
  TIME_ZONE_INFORMATION tzi;
  int nTZBias = 0;
  DWORD dwTZI = GetTimeZoneInformation(&tzi);
  if (dwTZI == TIME_ZONE_ID_DAYLIGHT)
    nTZBias = tzi.Bias + tzi.DaylightBias;
  else if (dwTZI != TIME_ZONE_ID_INVALID)
    nTZBias = tzi.Bias;
  char szTZBias[16];
  sprintf_s(szTZBias, _countof(szTZBias), "%+.2d%.2d", -nTZBias/60, abs(nTZBias)%60);

  //Get the local time
  SYSTEMTIME localTime;
  GetLocalTime(&localTime);

  //Validate the values in the SYSTEMTIME struct  
  ATLASSUME(localTime.wMonth >= 1 && localTime.wMonth <= 12);
  ATLASSUME(localTime.wDayOfWeek <= 6);

  //Finally form the "Date:" header
  char szDate[64];
  sprintf_s(szDate, _countof(szDate), "Date: %s, %d %s %04d %02d:%02d:%02d %s\r\n", pszDOW[localTime.wDayOfWeek], static_cast<int>(localTime.wDay), pszMonth[localTime.wMonth-1], static_cast<int>(localTime.wYear), 
            static_cast<int>(localTime.wHour), static_cast<int>(localTime.wMinute), static_cast<int>(localTime.wSecond), szTZBias);

  return szDate;
}

CPJNSMTPAsciiString CPJNSMTPMessage::GetHeader()
{
  CPJNSMTPString sCharset(m_RootPart.GetCharset());

  //Add the From field. Note we do not support encoding mail address headers for compatibility purposes
  CPJNSMTPAsciiString sFrom("From: ");
  sFrom += m_From.GetRegularFormat(m_bAddressHeaderEncoding, sCharset);
  CPJNSMTPAsciiString sHeader(sFrom);
  sHeader += "\r\n";
  
  //Add the "To:" field
  CPJNSMTPAsciiString sTo("To: ");
#ifdef CPJNSMTP_MFC_EXTENSIONS
  INT_PTR nTo = m_To.GetSize();
#else
  INT_PTR nTo = m_To.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<nTo; i++)
  {
    CPJNSMTPAsciiString sLine(m_To[i].GetRegularFormat(m_bAddressHeaderEncoding, sCharset));
    if (i < (nTo - 1))
      sLine += ",";
    sTo += sLine;
    if (i < (nTo - 1))
      sTo += "\r\n ";
  }
  sHeader += sTo;
  sHeader += "\r\n";

  //Create the "Cc:" part of the header
  CPJNSMTPAsciiString sCc("Cc: ");
#ifdef CPJNSMTP_MFC_EXTENSIONS
  INT_PTR nCC = m_CC.GetSize();
#else
  INT_PTR nCC = m_CC.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<nCC; i++)
  {
    CPJNSMTPAsciiString sLine(m_CC[i].GetRegularFormat(m_bAddressHeaderEncoding, sCharset));
    if (i < (nCC - 1))
      sLine += ",";
    sCc += sLine;
    if (i < (nCC - 1))
      sCc += "\r\n ";
  }

  //Add the CC field if there is any CC recipients
  if (nCC)
  {
    sHeader += sCc;
    sHeader += "\r\n";
  }

  //add the subject
  CPJNSMTPAsciiString sSubject(CPJNSMTPBodyPart::FoldSubjectHeader(m_sSubject, sCharset));
  sHeader += sSubject;
  sHeader += "\r\n";

  //X-Mailer header
#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (m_sXMailer.GetLength())
#else
  if (m_sXMailer.length())
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  { 
    CPJNSMTPAsciiString sXMailer("X-Mailer: ");
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    sXMailer += m_sXMailer;
  #else
  #ifdef _UNICODE
    sXMailer += ATL::CT2A(m_sXMailer.c_str());
  #else
    sXMailer += m_sXMailer;
  #endif //#ifdef _UNICODE
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    sHeader += sXMailer;
    sHeader += "\r\n";
  }
  
  //Disposition-Notification-To header
#ifdef CPJNSMTP_MFC_EXTENSIONS
  INT_PTR nMessageDispositionEmailAddresses = m_MessageDispositionEmailAddresses.GetSize();
#else
  INT_PTR nMessageDispositionEmailAddresses = m_MessageDispositionEmailAddresses.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (nMessageDispositionEmailAddresses)
  {
    CPJNSMTPAsciiString sMDN("Disposition-Notification-To: ");
    for (INT_PTR i=0; i<nMessageDispositionEmailAddresses; i++)
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      sMDN += m_MessageDispositionEmailAddresses[i];
    #else
    #ifdef _UNICODE
      sMDN += ATL::CT2A(m_MessageDispositionEmailAddresses[i].c_str());
    #else
      sMDN += m_MessageDispositionEmailAddresses[i];
    #endif //#ifdef _UNICODE
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      if (i < (nMessageDispositionEmailAddresses - 1))
        sMDN += ",\r\n ";
    }
    sHeader += sMDN;
    sHeader += "\r\n";
  }
  
  //Add the Mime header if needed
  if (m_bMime)
  {
    if (m_RootPart.GetNumberOfChildBodyParts())
    {
      CPJNSMTPAsciiString sPartHeader;
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      sPartHeader.Format("MIME-Version: 1.0\r\nContent-Type: %s; boundary=\"%s\"\r\n", CPJNSMTPAsciiString(m_RootPart.GetContentType()).operator LPCSTR(), CPJNSMTPAsciiString(m_RootPart.GetBoundary()).operator LPCSTR());
    #else
      std::ostringstream ss;
    #ifdef _UNICODE
      ss << "MIME-Version: 1.0\r\nContent-Type: " << ATL::CT2A(m_RootPart.GetContentType().c_str()) << "; boundary=\"" << ATL::CT2A(m_RootPart.GetBoundary().c_str()) << "\"\r\n";
    #else
      ss << "MIME-Version: 1.0\r\nContent-Type: " << m_RootPart.GetContentType() << "; boundary=\"" << m_RootPart.GetBoundary() << "\"\r\n";
    #endif //#ifdef _UNICODE
      sPartHeader = ss.str();
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      sHeader += sPartHeader;
    }
    else
    {
      CPJNSMTPAsciiString sPartHeader;
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      sPartHeader.Format("MIME-Version: 1.0\r\nContent-Type: %s\r\n", CPJNSMTPAsciiString(m_RootPart.GetContentType()).operator LPCSTR());
    #else
      std::ostringstream ss;
    #ifdef _UNICODE
      ss << "MIME-Version: 1.0\r\nContent-Type: " << ATL::CT2A(m_RootPart.GetContentType().c_str()) << "\r\n";
    #else
      ss << "MIME-Version: 1.0\r\nContent-Type: " << m_RootPart.GetContentType() << "\r\n";
    #endif //#ifdef _UNICODE
      sPartHeader = ss.str();
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      sHeader += sPartHeader;
    }
  }
  else
  {
    CPJNSMTPAsciiString sPartHeader;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    sPartHeader.Format("Content-Type: %s;\r\n\tcharset=%s\r\n", CPJNSMTPAsciiString(m_RootPart.GetContentType()).operator LPCSTR(), CPJNSMTPAsciiString(m_RootPart.GetCharset()).operator LPCSTR());
  #else
    std::ostringstream ss;
  #ifdef _UNICODE
    ss << "Content-Type: " << ATL::CT2A(m_RootPart.GetContentType().c_str()) << ";\r\n\tcharset=" << ATL::CT2A(m_RootPart.GetCharset().c_str()) << "\r\n";
  #else
    ss << "Content-Type: " << m_RootPart.GetContentType() << ";\r\n\tcharset=" << m_RootPart.GetCharset() << "\r\n";
  #endif //#ifdef _UNICODE
    sPartHeader = ss.str();
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    sHeader += sPartHeader;
  }
  
  //Add the optional Reply-To header if needed
#ifdef CPJNSMTP_MFC_EXTENSIONS
  INT_PTR nReplyTo = m_ReplyTo.GetSize();
#else
  INT_PTR nReplyTo = m_ReplyTo.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (nReplyTo)
  {
    CPJNSMTPAsciiString sReplyTo("Reply-To: ");
    for (INT_PTR i=0; i<nReplyTo; i++)
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      ATLASSERT(m_ReplyTo[i].m_sEmailAddress.GetLength());
      sReplyTo += m_ReplyTo[i].m_sEmailAddress;
    #else
      ATLASSERT(m_ReplyTo[i].m_sEmailAddress.length());
    #ifdef _UNICODE
      sReplyTo += ATL::CT2A(m_ReplyTo[i].m_sEmailAddress.c_str());
    #else
      sReplyTo += m_ReplyTo[i].m_sEmailAddress;
    #endif //#ifdef _UNICODE
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      if (i < (nReplyTo - 1))
        sReplyTo += ",\r\n ";
    }

    sHeader += sReplyTo;
    sHeader += "\r\n";
  }

  //Date header
  sHeader += FormDateHeader();

  //The Priorty header
  switch (m_Priority)
  {
    case NoPriority:
    {
      break;
    }
    case LowPriority:
    {
      sHeader += "X-Priority: 5\r\n";
      break;
    }
    case NormalPriority:
    {
      sHeader += "X-Priority: 3\r\n";
      break;
    }
    case HighPriority:
    {
      sHeader += "X-Priority: 1\r\n";
      break;
    }
    default:
    {
      ATLASSERT(FALSE);
      break;
    }
  }

  //Add the custom headers
#ifdef CPJNSMTP_MFC_EXTENSIONS
  INT_PTR nCustomHeaders = m_CustomHeaders.GetSize();
  for (INT_PTR i=0; i<nCustomHeaders; i++)
  {
    sHeader += m_CustomHeaders[i];
#else
  CPJNSMTPStringArray::size_type nCustomHeaders = m_CustomHeaders.size();
  for (CPJNSMTPStringArray::size_type i=0; i<nCustomHeaders; i++)
  {
  #ifdef _UNICODE
    sHeader += ATL::CT2A(m_CustomHeaders[i].c_str());
  #else
    sHeader += m_CustomHeaders[i];
  #endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    
    //Add line separators for each header
    sHeader += "\r\n";
  }

  //Return the result
  return sHeader;
}

INT_PTR CPJNSMTPMessage::ParseMultipleRecipients(_In_z_ LPCTSTR pszRecipients, _Inout_ CPJNSMTPMessage::CAddressArray& recipients)
{
  //Validate our parameters
  ATLASSERT(pszRecipients != NULL);

  //Empty out the array
#ifdef CPJNSMTP_MFC_EXTENSIONS
  recipients.SetSize(0);
#else
  recipients.clear();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  
  //Loop through the whole string, adding recipients as they are encountered
  CPJNSMTPString sRecipients(pszRecipients);
#ifdef CPJNSMTP_MFC_EXTENSIONS
  TCHAR* pszbuf = sRecipients.GetBuffer();
  int nLength = sRecipients.GetLength();
  ATLASSERT(nLength); //An empty string is now allowed
#else
  TCHAR* pszbuf = &(sRecipients[0]);
  int nLength = static_cast<int>(sRecipients.length());
  ATLASSERT(nLength); //An empty string is now allowed
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  BOOL bLeftQuotationMark = FALSE;
  for (int pos=0, start=0; pos<=nLength; pos++)
  {
    if (bLeftQuotationMark && pszbuf[pos] != _T('"') )
      continue;
    if (bLeftQuotationMark && pszbuf[pos] == _T('"'))
    {
      bLeftQuotationMark = FALSE;
      continue;
    }
    if (pszbuf[pos] == _T('"'))
    {
      bLeftQuotationMark = TRUE;
      continue;
    }
  
    //Valid separators between addresses are ',' or ';'
    if ((pszbuf[pos] == _T(',')) || (pszbuf[pos] == _T(';')) || (pszbuf[pos] == 0))
    {
      pszbuf[pos] = 0;	//Redundant when at the end of string, but who cares.
      CPJNSMTPString sTemp(&(pszbuf[start]));
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      sTemp.Trim();
    #else
      CPJNSMTPAddress::StringTrim(sTemp);
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

      //Let the CPJNSMTPAddress constructor do its work
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      CPJNSMTPAddress To(sTemp);
      if (To.m_sEmailAddress.GetLength())
        recipients.Add(To);
    #else
      CPJNSMTPAddress To(sTemp.c_str());
      if (To.m_sEmailAddress.length())
        recipients.push_back(To);
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

      //Move on to the next position
      start = pos + 1;
    }
  }

  //Return the number of recipients parsed
#ifdef CPJNSMTP_MFC_EXTENSIONS
  sRecipients.ReleaseBuffer();
  return recipients.GetSize();
#else
  return recipients.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

int CPJNSMTPMessage::AddMultipleAttachments(_In_z_ LPCTSTR pszAttachments)
{
  //Validate our parameters
  ATLASSERT(pszAttachments != NULL);

  //What will be the return value from this method
  int nAttachments = 0;

  //Loop through the whole string, adding attachments as they are encountered
  CPJNSMTPString sAttachments(pszAttachments);
#ifdef CPJNSMTP_MFC_EXTENSIONS
  TCHAR* pszbuf = sAttachments.GetBuffer();
  int nLength = sAttachments.GetLength();
  ATLASSERT(nLength); //An empty string is now allowed
#else
  TCHAR* pszbuf = &(sAttachments[0]);
  int nLength = static_cast<int>(sAttachments.length());
  ATLASSERT(nLength); //An empty string is now allowed
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  for (int pos=0, start=0; pos<=nLength; pos++)
  {
    //Valid separators between attachments are ',' or ';'
    if ((pszbuf[pos] == _T(',')) || (pszbuf[pos] == _T(';')) || (pszbuf[pos] == 0))
    {
      pszbuf[pos] = 0;	//Redundant when at the end of string, but who cares.
      CPJNSMTPString sTemp(&(pszbuf[start]));

      //Finally add the new attachment to the array of attachments
      CPJNSMTPBodyPart attachment;
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      sTemp.Trim();
      if (sTemp.GetLength())
    #else
      CPJNSMTPAddress::StringTrim(sTemp);
      if (sTemp.length())
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      {
        BOOL bAdded = attachment.SetFilename(sTemp);
        if (bAdded)
        {
          ++nAttachments;
          AddBodyPart(attachment);
        }
      }

      //Move on to the next position
      start = pos + 1;
    }
  }

#ifdef CPJNSMTP_MFC_EXTENSIONS
  sAttachments.ReleaseBuffer();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  return nAttachments;
}

void CPJNSMTPMessage::AddTextBody(_In_z_ LPCTSTR pszBody, _In_z_ LPCTSTR pszRootMIMEType)
{
  if (m_bMime)
  {
    CPJNSMTPBodyPart* pTextBodyPart = m_RootPart.FindFirstBodyPart(_T("text/plain"));
    if (pTextBodyPart != NULL)
      pTextBodyPart->SetText(pszBody);
    else
    {
      //Create a text body part
      CPJNSMTPBodyPart oldRoot(m_RootPart);

      //Reset the root body part to be multipart/mixed
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      m_RootPart.SetCharset(oldRoot.GetCharset());
    #else
      m_RootPart.SetCharset(oldRoot.GetCharset().c_str());
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      m_RootPart.SetText(_T("This is a multi-part message in MIME format"));
      m_RootPart.SetContentType(pszRootMIMEType);

      //Just add the text/plain body part (directly to the root)
      CPJNSMTPBodyPart text;
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      text.SetCharset(oldRoot.GetCharset());
    #else
      text.SetCharset(oldRoot.GetCharset().c_str());
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      text.SetText(pszBody);

      //Hook everything up to the root body part
      m_RootPart.InsertChildBodyPart(text);
    }
  }
  else
  {
    //Let the body part class do all the work
    m_RootPart.SetText(pszBody);
  }
}

void CPJNSMTPMessage::AddHTMLBody(_In_z_ LPCTSTR pszBody, _In_z_ LPCTSTR pszRootMIMEType)
{
  if (m_bMime)
  {
    CPJNSMTPBodyPart* pHtmlBodyPart = m_RootPart.FindFirstBodyPart(_T("text/html"));
    if (pHtmlBodyPart != NULL)
      pHtmlBodyPart->SetText(pszBody);
    else
    {
      //Remember some of the old root settings before we write over it
      CPJNSMTPBodyPart oldRoot(m_RootPart);

      //Reset the root body part to be multipart/mixed
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      m_RootPart.SetCharset(oldRoot.GetCharset());
    #else
      m_RootPart.SetCharset(oldRoot.GetCharset().c_str());
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      m_RootPart.SetText(_T("This is a multi-part message in MIME format"));
      m_RootPart.SetContentType(pszRootMIMEType);

      //Just add the text/html body part (directly to the root)
      CPJNSMTPBodyPart html;
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      html.SetCharset(oldRoot.GetCharset());
    #else
      html.SetCharset(oldRoot.GetCharset().c_str());
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      html.SetText(pszBody);
      html.SetContentType(_T("text/html"));

      //Hook everything up to the root body part
      m_RootPart.InsertChildBodyPart(html);
    }
  }
  else
  {
    m_RootPart.SetText(pszBody);
    m_RootPart.SetContentType(_T("text/html"));
  }
}

CPJNSMTPString CPJNSMTPMessage::GetHTMLBody()
{
  //What will be the return value from this method
  CPJNSMTPString sRet;

  if (m_RootPart.GetNumberOfChildBodyParts())
  {
    CPJNSMTPBodyPart* pHtml = m_RootPart.GetChildBodyPart(0);
    if (pHtml->GetContentType() == _T("text/html"))
      sRet = pHtml->GetText();
  }
  return sRet;
}

CPJNSMTPString CPJNSMTPMessage::GetTextBody()
{
  return m_RootPart.GetText();
}

void CPJNSMTPMessage::SetMime(BOOL bMime)
{
  if (m_bMime != bMime)
  {
    m_bMime = bMime;

    //Reset the body body parts
    for (INT_PTR i=0; i<m_RootPart.GetNumberOfChildBodyParts(); i++)
      m_RootPart.RemoveChildBodyPart(i);

    if (bMime)
    {
      CPJNSMTPString sText(GetTextBody());

      //Remember some of the old root settings before we write over it
      CPJNSMTPBodyPart oldRoot(m_RootPart);

      //Reset the root body part to be multipart/mixed
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      m_RootPart.SetCharset(oldRoot.GetCharset());
    #else
      m_RootPart.SetCharset(oldRoot.GetCharset().c_str());
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      m_RootPart.SetText(_T("This is a multi-part message in MIME format"));
      m_RootPart.SetContentType(_T("multipart/mixed"));

      //Also readd the body if non - empty
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      if (sText.GetLength())
        AddTextBody(sText);
    #else
      if (sText.length())
        AddTextBody(sText.c_str());
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
  }
}

void CPJNSMTPMessage::SaveToDisk(_In_z_ LPCTSTR pszFilename)
{
  //Open the file for writing
  ATL::CAtlFile file;
  HRESULT hr = file.Create(pszFilename, GENERIC_WRITE, 0, OPEN_ALWAYS);
  if (FAILED(hr))
    CPJNSMTPConnection::ThrowPJNSMTPException(hr);

  //Set the file back to 0 in size
  hr = file.SetSize(0);
  if (FAILED(hr))
    CPJNSMTPConnection::ThrowPJNSMTPException(hr);

  //Write out the Message Header
  CPJNSMTPAsciiString sHeader(GetHeader());
#ifdef CPJNSMTP_MFC_EXTENSIONS
  hr = file.Write(sHeader.operator LPCSTR(), sHeader.GetLength());
#else
  hr = file.Write(sHeader.c_str(), static_cast<DWORD>(sHeader.length()));
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (FAILED(hr))
    CPJNSMTPConnection::ThrowPJNSMTPException(hr);

  //Write out the separator
  hr = file.Write("\r\n", 2);
  if (FAILED(hr))
    CPJNSMTPConnection::ThrowPJNSMTPException(hr);

  //Write out the rest of the message
  if (m_RootPart.GetNumberOfChildBodyParts() || m_bMime)
  {
    //Write the root body part (and all its children)
    WriteToDisk(file, &m_RootPart, TRUE, m_RootPart.GetCharset());
  }
  else
  {
    //Send the body
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    CPJNSMTPAsciiString sBody(m_RootPart.GetText());
    hr = file.Write(sBody.operator LPCSTR(), sBody.GetLength());
  #else

  #ifdef _UNICODE
    CPJNSMTPAsciiString sBody(ATL::CT2A(m_RootPart.GetText().c_str()));
  #else
    CPJNSMTPAsciiString sBody(m_RootPart.GetText());
  #endif //#ifdef _UNICODE
    hr = file.Write(sBody.c_str(), static_cast<DWORD>(sBody.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    if (FAILED(hr))
      CPJNSMTPConnection::ThrowPJNSMTPException(hr);
  }
}

void CPJNSMTPMessage::WriteToDisk(ATL::CAtlFile& file, CPJNSMTPBodyPart* pBodyPart, BOOL bRoot, const CPJNSMTPString& sRootCharset)
{
  //Validate our parameters
  ATLASSUME(pBodyPart != NULL);

  if (!bRoot)
  {
    //First send this body parts header
    CPJNSMTPAsciiString sHeader(pBodyPart->GetHeader(sRootCharset));
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    HRESULT hr = file.Write(sHeader.operator LPCSTR(), sHeader.GetLength());
  #else
    HRESULT hr = file.Write(sHeader.c_str(), static_cast<DWORD>(sHeader.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    if (FAILED(hr))
      CPJNSMTPConnection::ThrowPJNSMTPException(hr);
  }
  
  //Then the body parts body
  CPJNSMTPAsciiString sBody(pBodyPart->GetBody(FALSE));
#ifdef CPJNSMTP_MFC_EXTENSIONS
  HRESULT hr = file.Write(sBody.operator LPCSTR(), sBody.GetLength());
#else
  HRESULT hr = file.Write(sBody.c_str(), static_cast<DWORD>(sBody.length()));
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (FAILED(hr))
    CPJNSMTPConnection::ThrowPJNSMTPException(hr);

  //Recursively save all the child body parts
  INT_PTR nChildBodyParts = pBodyPart->GetNumberOfChildBodyParts();
  for (INT_PTR i=0; i<nChildBodyParts; i++)
  {
    CPJNSMTPBodyPart* pChildBodyPart = pBodyPart->GetChildBodyPart(i);
    WriteToDisk(file, pChildBodyPart, FALSE, sRootCharset);
  }

  //Then the MIME footer if need be
  BOOL bSendFooter = (pBodyPart->GetNumberOfChildBodyParts() != 0);
  if (bSendFooter)
  {
    CPJNSMTPAsciiString sFooter(pBodyPart->GetFooter());
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    hr = file.Write(sFooter.operator LPCSTR(), sFooter.GetLength());
  #else
    hr = file.Write(sFooter.c_str(), static_cast<DWORD>(sFooter.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    if (FAILED(hr))
      CPJNSMTPConnection::ThrowPJNSMTPException(hr);
  }
}


CPJNSMTPConnection::CPJNSMTPConnection() : m_bConnected(FALSE), 
                                           m_nLastCommandResponseCode(0), 
                                           m_ConnectionType(PlainText),
                                           m_bCanDoDSN(FALSE),
                                           m_bCanDoSTARTTLS(FALSE),
                                           m_bDoInitialClientResponse(FALSE),
                                           m_SSLProtocol(OSDefault),
                                           m_dwTimeout(60000)
{
  SetHeloHostname(_T("auto"));
}

CPJNSMTPConnection::~CPJNSMTPConnection()
{
  if (m_bConnected)
  {
    //Deliberately handle exceptions here, so that the destructor does not throw exception
    try
    {
      Disconnect(TRUE);
    }
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    catch(CPJNSMTPException* pEx)
    {
      pEx->Delete();
    }
  #else
    catch(CPJNSMTPException& /*e*/)
    {
    }
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
}

void CPJNSMTPConnection::ThrowPJNSMTPException(_In_ DWORD dwError, _In_ DWORD dwFacility, _In_ const CPJNSMTPString& sLastResponse)
{
  if (dwError == 0)
    dwError = ::WSAGetLastError();

#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPException* pException = new CPJNSMTPException(dwError, dwFacility, sLastResponse);
  TRACE(_T("Warning: throwing CPJNSMTPException for error %08X\n"), pException->m_hr);
  THROW(pException);
#else
  CPJNSMTPException e(dwError, dwFacility, sLastResponse);
  ATLTRACE(_T("Warning: throwing CPJNSMTPException for error %08X\n"), e.m_hr);
  throw e;
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

void CPJNSMTPConnection::ThrowPJNSMTPException(_In_ HRESULT hr, _In_ const CPJNSMTPString& sLastResponse)
{
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPException* pException = new CPJNSMTPException(hr, sLastResponse);
  TRACE(_T("Warning: throwing CPJNSMTPException for error %08X\n"), pException->m_hr);
  THROW(pException);
#else
  CPJNSMTPException e(hr, sLastResponse);
  ATLTRACE(_T("Warning: throwing CPJNSMTPException for error %08X\n"), e.m_hr);
  throw e;
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
}

void CPJNSMTPConnection::SetHeloHostname(_In_z_ LPCTSTR pszHostname)
{
  //retrieve the localhost name if we are supplied the magic "auto" host name
  if (_tcscmp(pszHostname, _T("auto")) == 0)
  {
    //retrieve the localhost name
    char szHostName[_MAX_PATH];
    szHostName[0] = '\0';
    if (gethostname(szHostName, _countof(szHostName)) == 0)
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      m_sHeloHostname = szHostName;
    #else
    #ifdef _UNICODE
      m_sHeloHostname = ATL::CA2T(szHostName);
    #else
      m_sHeloHostname = szHostName;
    #endif //#ifdef _UNICODE
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
    else
      m_sHeloHostname = pszHostname;
  }
  else
    m_sHeloHostname = pszHostname;
}

void CPJNSMTPConnection::_Send(const void* pBuffer, int nBuf)
{
#ifndef CPJNSMTP_NOSSL
  if (GetSocket() != INVALID_SOCKET)
  {
    SECURITY_STATUS ss = SendEncrypted(static_cast<const BYTE*>(pBuffer), nBuf);
    if (ss != SEC_E_OK)
      ThrowPJNSMTPException(ss);
  }
  else 
#endif //#ifndef CPJNSMTP_NOSSL
    m_Socket.Send(pBuffer, nBuf);
}

#ifdef CPJNSMTP_NOSSL
BOOL CPJNSMTPConnection::DoEHLO(_In_ AuthenticationMethod am, _In_ ConnectionType /*connectionType*/, _In_ DWORD dwDSN)
#else
BOOL CPJNSMTPConnection::DoEHLO(_In_ AuthenticationMethod am, _In_ ConnectionType connectionType, _In_ DWORD dwDSN)
#endif //#ifdef CPJNSMTP_NOSSL
{
  //What will be the return value from this method (assume the worst)
  BOOL bDoEHLO = FALSE;

  //Determine if we should connect using EHLO instead of HELO
#ifndef CPJNSMTP_NOSSL
  if (connectionType == STARTTLS || connectionType == AutoUpgradeToSTARTTLS)
    bDoEHLO = TRUE;
#endif //#ifdef CPJNSMTP_NOSSL
  if (!bDoEHLO)
    bDoEHLO = (am != AUTH_NONE);
  if (!bDoEHLO)
    bDoEHLO = (dwDSN != CPJNSMTPMessage::DSN_NOT_SPECIFIED);

  return bDoEHLO;
}

#ifndef CPJNSMTP_NOSSL
void CPJNSMTPConnection::CreateSSLSocket()
{
  //Acquire the credentials
  if (m_SSLCredentials.ValidHandle())
    m_SSLCredentials.Free();
  switch (m_SSLProtocol)
  {
    case SSLv2:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = SP_PROT_SSL2_CLIENT;
       break;
    }
    case SSLv3:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = SP_PROT_SSL3_CLIENT;
       break;
    }
    case SSLv2orv3:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = SP_PROT_SSL2_CLIENT | SP_PROT_SSL3_CLIENT;
       break;
    }
    case TLSv1:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = SP_PROT_TLS1_0_CLIENT;
       break;
    }
    case TLSv1_1:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = SP_PROT_TLS1_1_CLIENT;
       break;
    }
    case TLSv1_2:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = SP_PROT_TLS1_2_CLIENT;
       break;
    }
    case OSDefault:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = 0;
       break;
    }
    case AnyTLS:
    {
       m_SSLCredentials.m_sslCredentials.cSupportedAlgs = SP_PROT_TLS1_0_CLIENT | SP_PROT_TLS1_1_CLIENT | SP_PROT_TLS1_2_CLIENT;
       break;
    }
    default:
    {
      ATLASSERT(FALSE);
      break;
    }
  }

  SECURITY_STATUS ss = m_SSLCredentials.AcquireClient();
  if (ss != SEC_E_OK)
    ThrowPJNSMTPException(ss);

  //Setup the SSLWrappers::CSocket instance
  if (ValidHandle())
    Delete();
  SetCachedCredentials(&m_SSLCredentials);
  SetReadTimeout(m_dwTimeout);
  SetWriteTimeout(m_dwTimeout);
  Attach(m_Socket);
}
#endif //#ifdef CPJNSMTP_NOSSL

void CPJNSMTPConnection::Connect(_In_z_ LPCTSTR pszHostName, _In_ AuthenticationMethod am, _In_opt_z_ LPCTSTR pszUsername, _In_opt_z_ LPCTSTR pszPassword, _In_ int nPort, _In_ ConnectionType connectionType, _In_ DWORD dwDSN)
{
  //Validate our parameters
  ATLASSERT(pszHostName != NULL);
  ATLASSERT(!m_bConnected);

#ifndef CPJNSMTP_NOSSL
  m_ConnectionType = connectionType;
#else
  m_ConnectionType = PlainText;
#endif //#ifndef CPJNSMTP_NOSSL
  m_sHostName = pszHostName;

  try
  {
#ifdef CWSOCKET_MFC_EXTENSIONS
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    m_Socket.SetBindAddress(m_sBindAddress);
  #else
    m_Socket.SetBindAddress(CWSocket::String(m_sBindAddress.c_str()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
#else
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    m_Socket.SetBindAddress(CWSocket::String(m_sBindAddress));
  #else
    m_Socket.SetBindAddress(m_sBindAddress);
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS
    m_Socket.CreateAndConnect(pszHostName, nPort, SOCK_STREAM, AF_INET);

  #ifndef CPJNSMTP_NOSSL
    if (m_ConnectionType == SSL_TLS)
    {
      CreateSSLSocket();
      SECURITY_STATUS ss = SSLConnect(pszHostName);
      if (ss != SEC_E_OK)
      {
        Detach();

        //Disconnect before we throw the exception
        try
        {
          Disconnect(FALSE);
        }
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        catch(CPJNSMTPException* pEx2)
        {
          pEx2->Delete();
        }
      #else
        catch(CPJNSMTPException& /*e*/)
        {
        }
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
        ThrowPJNSMTPException(ss);
      }
    }
  #endif //#ifdef CPJNSMTP_NOSSL
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
#else
  catch(CWSocketException& e)
  {
    DWORD dwError = e.m_nError;
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

    //Disconnect before we rethrow the exception
    try
    {
      Disconnect(TRUE);
    }
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    catch(CPJNSMTPException* pEx2)
    {
      pEx2->Delete();
    }
  #else
    catch(CPJNSMTPException& /*e*/)
    {
    }
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

    //rethrow the exception
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }

  //We're now connected
  m_bConnected = TRUE;

  try
  {
    //check the response to the login
    if (!ReadCommandResponse(220))
      ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_SMTP_LOGIN_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

    //Negotiate Extended SMTP connection if required
    if (DoEHLO(am, connectionType, dwDSN))
    {
      //ATLASSUME(pszUsername != NULL);
      //ATLASSUME(pszPassword != NULL);

    #ifdef CPJNSMTP_MFC_EXTENSIONS
      ConnectESMTP(m_sHeloHostname, pszUsername, pszPassword, am);
    #else
      ConnectESMTP(m_sHeloHostname.c_str(), pszUsername, pszPassword, am);
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
    else
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      ConnectSMTP(m_sHeloHostname);
    #else
      ConnectSMTP(m_sHeloHostname.c_str());
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
  }
#ifdef CPJNSMTP_MFC_EXTENSIONS
  catch(CPJNSMTPException* pEx)
  {
#else
  catch(CPJNSMTPException& e)
  {
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    //Disconnect before we rethrow the exception
    try
    {
      Disconnect(TRUE);
    }
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    catch(CPJNSMTPException* pEx2)
    {
      pEx2->Delete();
    }
  #else
    catch(CPJNSMTPException& /*e*/)
    {
    }
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  #ifdef CPJNSMTP_MFC_EXTENSIONS
    HRESULT hr = pEx->m_hr;
    CPJNSMTPString sLastResponse(pEx->m_sLastResponse);
    pEx->Delete();
  #else
    HRESULT hr = e.m_hr;
    CPJNSMTPString sLastResponse(e.m_sLastResponse);
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

    //rethrow the exception
    ThrowPJNSMTPException(hr, sLastResponse);
  }
}

CPJNSMTPConnection::AuthenticationMethod CPJNSMTPConnection::ChooseAuthenticationMethod(_In_z_ LPCTSTR pszAuthMethods)
{
  //What will be the return value from this method
  AuthenticationMethod am = AUTH_AUTO;

  //Decide on which protocol to choose. Note we deliberately do not auto pick XOAUTH2 even if it is supported because
  //it requires some additional ahead of time setup work. If you really want to use XOAUTH2 authentication, then just
  //explicitly pick it as the authentication mechanism
  if (_tcsstr(pszAuthMethods, _T("NTLM")) != NULL)
    am = AUTH_NTLM;
  else if (_tcsstr(pszAuthMethods, _T("CRAM-MD5")) != NULL)
    am = AUTH_CRAM_MD5;
  else if (_tcsstr(pszAuthMethods, _T("LOGIN")) != NULL)
    am = AUTH_LOGIN;
  else if (_tcsstr(pszAuthMethods, _T("PLAIN")) != NULL)
    am = AUTH_PLAIN;
  else 
    am = AUTH_NONE;
    
  return am;
}

//This method implmements the STARTTLS functionality
#ifndef CPJNSMTP_NOSSL
void CPJNSMTPConnection::DoSTARTTLS()
{
  static char szStart[] = "STARTTLS\r\n";
 
  //Lets issue the STARTTLS command
  try
  {
    _Send(szStart, static_cast<int>(strlen(szStart)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS
 
  //Check the response
  if (!ReadCommandResponse(220))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_STARTTLS_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
 
  //Hook up the SSL socket at this point
  CreateSSLSocket();
 
  //Do the SSL connect
#ifdef CPJNSMTP_MFC_EXTENSIONS
  SECURITY_STATUS ss = SSLConnect(m_sHostName);
#else
  SECURITY_STATUS ss = SSLConnect(m_sHostName.c_str());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (ss != SEC_E_OK)
  {
    Detach();
    ThrowPJNSMTPException(ss);
  }
}
#endif //#ifdef CPJNSMTP_NOSSL

void CPJNSMTPConnection::SendEHLO(_In_z_ LPCTSTR pszLocalName, _Inout_ CPJNSMTPString& sLastResponse)
{
  //Send the EHLO command
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sBuf;
  sBuf.Format("EHLO %s\r\n", CPJNSMTPAsciiString(pszLocalName).operator LPCSTR());
  try
  {
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
  }
#else
  CPJNSMTPAsciiString sBuf;
  std::ostringstream ss;
#ifdef _UNICODE
  ss << "EHLO " << ATL::CT2A(pszLocalName) << "\r\n";
#else
  ss << "EHLO " << pszLocalName << "\r\n";
#endif //#ifdef _UNICODE
  sBuf = ss.str();
  try
  {
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
  }
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS
  
  //check the response to the EHLO command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_EHLO_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  sLastResponse = GetLastCommandResponse();
}

//This method connects using RFC1869 ESMTP 
void CPJNSMTPConnection::ConnectESMTP(_In_z_ LPCTSTR pszLocalName, 
                                      _When_(am == AUTH_CRAM_MD5, _In_z_)
                                      _When_(am == AUTH_LOGIN, _In_z_)
                                      _When_(am == AUTH_PLAIN, _In_z_)
                                      _When_(am == AUTH_XOAUTH2, _In_z_) 
                                      _When_((am != AUTH_CRAM_MD5) && (am != AUTH_LOGIN) && (am != AUTH_PLAIN) && (am != AUTH_XOAUTH2), _In_opt_z_) 
                                      LPCTSTR pszUsername, 
                                      _When_(am == AUTH_CRAM_MD5, _In_z_)
                                      _When_(am == AUTH_LOGIN, _In_z_)
                                      _When_(am == AUTH_PLAIN, _In_z_)
                                      _When_(am == AUTH_XOAUTH2, _In_z_) 
                                      _When_((am != AUTH_CRAM_MD5) && (am != AUTH_LOGIN) && (am != AUTH_PLAIN) && (am != AUTH_XOAUTH2), _In_opt_z_) 
                                      LPCTSTR pszPassword, _In_ AuthenticationMethod am)
{
  //Send the EHLO command
  CPJNSMTPString sResponse;
  SendEHLO(pszLocalName, sResponse);
  CPJNSMTPString sUpperCaseResponse(sResponse);
#ifdef CPJNSMTP_MFC_EXTENSIONS
  sUpperCaseResponse.MakeUpper();

  //Check the server capabilities
  m_bCanDoDSN = (sUpperCaseResponse.Find(_T("250-DSN\r\n")) >= 0) || (sUpperCaseResponse.Find(_T("250 DSN\r\n")) >= 0);
  m_bCanDoSTARTTLS = (sUpperCaseResponse.Find(_T("250-STARTTLS\r\n")) >= 0) || (sUpperCaseResponse.Find(_T("250 STARTTLS\r\n")) >= 0);
#else
  std::transform(sUpperCaseResponse.begin(), sUpperCaseResponse.end(), sUpperCaseResponse.begin(), _totupper);

  //Check the server capabilities
  m_bCanDoDSN = (sUpperCaseResponse.find(_T("250-DSN\r\n")) != CPJNSMTPString::npos) || (sUpperCaseResponse.find(_T("250 DSN\r\n")) != CPJNSMTPString::npos);
  m_bCanDoSTARTTLS = (sUpperCaseResponse.find(_T("250-STARTTLS\r\n")) != CPJNSMTPString::npos) || (sUpperCaseResponse.find(_T("250 STARTTLS\r\n")) != CPJNSMTPString::npos);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  
  //Check to see if we should do STARTTLS
#ifndef CPJNSMTP_NOSSL
  BOOL bDoSTARTTLS = FALSE;
  switch (m_ConnectionType)
  {
    case STARTTLS:
    {
      //Yes we will be explicitly doing STARTTLS
      bDoSTARTTLS = TRUE;
      break;
    }
    case AutoUpgradeToSTARTTLS:
    {
      //Check the response to see if the server supports STARTTLS
      bDoSTARTTLS = m_bCanDoSTARTTLS;
      break;
    }
    default:
    {
      break;
    }
  }

  //Do STARTTLS if required to do so
  if (bDoSTARTTLS)
  {
    DoSTARTTLS();

    //Reissue the EHLO command to ensure we see the most up to date response values after we have done a STARTTLS
    SendEHLO(pszLocalName, sResponse);
    sUpperCaseResponse = sResponse;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    sUpperCaseResponse.MakeUpper();
  #else
    std::transform(sUpperCaseResponse.begin(), sUpperCaseResponse.end(), sUpperCaseResponse.begin(), _totupper);
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#endif //#ifdef CPJNSMTP_NOSSL

  //If we are doing Auto authentication, then find the supported protocols of the SMTP server, then pick and use one
  if (am == AUTH_AUTO)
  {
    //Find the authentication methods supported by the SMTP server
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    int nPosBegin = sUpperCaseResponse.Find(_T("250-AUTH "));
    if (nPosBegin == -1)
      nPosBegin = sUpperCaseResponse.Find(_T("250 AUTH "));
    if (nPosBegin >= 0)
    {
      CPJNSMTPString sAuthMethods;
      int nPosEnd = sUpperCaseResponse.Find(_T("\r\n"), nPosBegin);
      if (nPosEnd >= 0)
        sAuthMethods = sUpperCaseResponse.Mid(nPosBegin + 9, nPosEnd - nPosBegin - 9);
      else
        sAuthMethods = sUpperCaseResponse.Mid(nPosBegin + 9);

      //Now decide which protocol to choose via the virtual function to allow further client customization
      am = ChooseAuthenticationMethod(sAuthMethods);
  #else
    CPJNSMTPString::size_type nPosBegin = sUpperCaseResponse.find(_T("250-AUTH "));
    if (nPosBegin == CPJNSMTPString::npos)
      nPosBegin = sUpperCaseResponse.find(_T("250 AUTH "));
    if (nPosBegin != CPJNSMTPString::npos)
    {
      CPJNSMTPString sAuthMethods;
      CPJNSMTPString::size_type nPosEnd = sUpperCaseResponse.find(_T("\r\n"), nPosBegin);
      if (nPosEnd != CPJNSMTPString::npos)
        sAuthMethods = sUpperCaseResponse.substr(nPosBegin + 9, nPosEnd - nPosBegin - 9);
      else
        sAuthMethods = sUpperCaseResponse.substr(nPosBegin + 9);

      //Now decide which protocol to choose via the virtual function to allow further client customization
      am = ChooseAuthenticationMethod(sAuthMethods.c_str());
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
    else
    {
      //If AUTH is not found the server does not support authentication
      am = AUTH_NONE;
    }
  }

  //What we do next depends on what authentication protocol we are using
  switch (am)
  {
    case AUTH_CRAM_MD5:
    {
      ATLASSUME(pszUsername != NULL);
      ATLASSUME(pszPassword != NULL);

      AuthCramMD5(pszUsername, pszPassword);
      break;
    }
    case AUTH_LOGIN:
    {
      ATLASSUME(pszUsername != NULL);
      ATLASSUME(pszPassword != NULL);

      AuthLogin(pszUsername, pszPassword);
      break;
    }
    case AUTH_PLAIN:
    {
      ATLASSUME(pszUsername != NULL);
      ATLASSUME(pszPassword != NULL);

      AuthPlain(pszUsername, pszPassword);
      break;
    }
    case AUTH_XOAUTH2:
    {
      ATLASSUME(pszUsername != NULL);
      ATLASSUME(pszPassword != NULL);

      AuthXOAUTH2(pszUsername, pszPassword);
      break;
    }
  #ifndef CPJNSMTP_NONTLM
    case AUTH_NTLM:
    {
      SECURITY_STATUS ss = NTLMAuthenticate(pszUsername, pszPassword);
      if (!SEC_SUCCESS(ss))
        ThrowPJNSMTPException(ss, m_sLastCommandResponse);
      break;
    }
  #endif //#ifdef CPJNSMTP_NONTLM
    case AUTH_AUTO:
    {
      ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_EHLO_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
      break;
    }
    case AUTH_NONE:
    {
      //authentication not requried
      break;    
    }
    default:
    {
      ATLASSERT(FALSE);
      break;
    }
  }
}

//This method connects using standard RFC 821 SMTP
void CPJNSMTPConnection::ConnectSMTP(_In_z_ LPCTSTR pszLocalName)
{
  //The server does not support any extentions if we use standard SMTP
  m_bCanDoDSN = FALSE;
  m_bCanDoSTARTTLS = FALSE;

  CPJNSMTPAsciiString sBuf;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  sBuf.Format("HELO %s\r\n", CPJNSMTPAsciiString(pszLocalName).operator LPCSTR());
  try
  {
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
#else
  std::ostringstream ss;
#ifdef _UNICODE
  ss << "HELO " << ATL::CT2A(pszLocalName) << "\r\n";
#else
  ss << "HELO " << pszLocalName << "\r\n";
#endif //#ifdef _UNICODE
  sBuf = ss.str();
  try
  {
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the HELO command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_HELO_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

void CPJNSMTPConnection::Disconnect(_In_ BOOL bGracefully)
{
  //disconnect from the SMTP server if connected 
  HRESULT hr = S_OK;
  CPJNSMTPString sLastResponse;
  if (m_bConnected)
  {
    if (bGracefully)
    {
      static char* sBuf = "QUIT\r\n";
      try
      {
        _Send(sBuf, static_cast<int>(strlen(sBuf)));
      }
    #ifdef CWSOCKET_MFC_EXTENSIONS
      catch(CWSocketException* pEx)
      {
        hr = MAKE_HRESULT(SEVERITY_ERROR, pEx->m_nError, FACILITY_WIN32);
        pEx->Delete();
      }
    #else
      catch(CWSocketException& e)
      {
        hr = MAKE_HRESULT(SEVERITY_ERROR, e.m_nError, FACILITY_WIN32);
      }
    #endif //#ifdef CWSOCKET_MFC_EXTENSIONS

      //Check the reponse
      try
      {
        if (!ReadCommandResponse(221))
        {
          hr = MAKE_HRESULT(SEVERITY_ERROR, IDS_PJNSMTP_UNEXPECTED_QUIT_RESPONSE, FACILITY_ITF);
          sLastResponse = m_sLastCommandResponse;
        }
      }
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      catch(CPJNSMTPException* pEx)
      {
        //Hive away the values in the local variables first. We do this to ensure that we 
        //end up executing the m_bConnected = FALSE and _Close code below
        hr = pEx->m_hr;
        sLastResponse = pEx->m_sLastResponse;
      
        pEx->Delete();
      }
    #else
      catch(CPJNSMTPException& e)
      {
        //Hive away the values in the local variables first. We do this to ensure that we 
        //end up executing the m_bConnected = FALSE and _Close code below
        hr = e.m_hr;
        sLastResponse = e.m_sLastResponse;
      }
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }

    //Reset the connected state
    m_bConnected = FALSE;
  }
 
  //free up our socket
#ifndef CPJNSMTP_NOSSL
  if (m_SSLCredentials.ValidHandle())
    m_SSLCredentials.Free();
  if (ValidHandle())
    Delete();
  Detach();
#endif //#ifdef CPJNSMTP_NOSSL
  m_Socket.Close();

  //Should we throw an expection
  if (FAILED(hr))
    ThrowPJNSMTPException(hr, sLastResponse);
}

void CPJNSMTPConnection::SendBodyPart(_In_ CPJNSMTPBodyPart* pBodyPart, _In_ BOOL bRoot, _In_z_ LPCTSTR pszRootCharset)
{
  //Validate our parameters
  ATLASSUME(pBodyPart != NULL);

  if (!bRoot)
  {
    //First send this body parts header
    CPJNSMTPAsciiString sHeader(pBodyPart->GetHeader(pszRootCharset));
    try
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      int nHeaderLength = sHeader.GetLength();
      if (nHeaderLength)
        _Send(sHeader.operator LPCSTR(), nHeaderLength);
    #else
      CPJNSMTPAsciiString::size_type nHeaderLength = sHeader.length();
      if (nHeaderLength)
        _Send(sHeader.c_str(), static_cast<int>(nHeaderLength));
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS
  }
  
  //Then the body parts body
  CPJNSMTPAsciiString sBody(pBodyPart->GetBody(TRUE));
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    int nBodyLength = sBody.GetLength();
    if (nBodyLength)
      _Send(sBody.operator LPCSTR(), nBodyLength);
  #else
    CPJNSMTPAsciiString::size_type nBodyLength = sBody.length();
    if (nBodyLength)
      _Send(sBody.c_str(), static_cast<int>(nBodyLength));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //The recursively send all the child body parts
  INT_PTR nChildBodyParts = pBodyPart->GetNumberOfChildBodyParts();
  for (INT_PTR i=0; i<nChildBodyParts; i++)
  {
    CPJNSMTPBodyPart* pChildBodyPart = pBodyPart->GetChildBodyPart(i);
    SendBodyPart(pChildBodyPart, FALSE, pszRootCharset);
  }

  //Then the MIME footer if need be
  BOOL bSendFooter = (pBodyPart->GetNumberOfChildBodyParts() != 0);
  if (bSendFooter)
  {
    CPJNSMTPAsciiString sFooter(pBodyPart->GetFooter());
    try
    {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
      int nFooterLength = sFooter.GetLength();
      if (nFooterLength)
        _Send(sFooter.operator LPCSTR(), nFooterLength);
  #else
      CPJNSMTPAsciiString::size_type nFooterLength = sFooter.length();
      if (nFooterLength)
        _Send(sFooter.c_str(), static_cast<int>(nFooterLength));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS 
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS
  }
}

CPJNSMTPAsciiString CPJNSMTPConnection::FormMailFromCommand(_In_z_ LPCTSTR pszEmailAddress, _In_ DWORD dwDSN, _In_ CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType, _Inout_ CPJNSMTPAsciiString& sENVID)
{
  //Validate our parameters
  ATLASSERT(pszEmailAddress != NULL);
  ATLASSERT(_tcslen(pszEmailAddress));
  
  //What will be the return value from this method
  CPJNSMTPAsciiString sBuf;
  
  if (dwDSN == CPJNSMTPMessage::DSN_NOT_SPECIFIED)
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS 
    sBuf.Format("MAIL FROM:<%s>\r\n", CPJNSMTPAsciiString(pszEmailAddress).operator LPCSTR());
  #else
    std::ostringstream ss;
  #ifdef _UNICODE
    ss << "MAIL FROM:<" << ATL::CT2A(pszEmailAddress) << ">\r\n";
  #else
    ss << "MAIL FROM:<" << pszEmailAddress << ">\r\n";
  #endif //#ifdef _UNICODE
    sBuf = ss.str();
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS 
  }
  else
  {
    //Fail if we have requested DSN and the server does not support it
    if (!m_bCanDoDSN)
      ThrowPJNSMTPException(IDS_PJNSMTP_DSN_NOT_SUPPORTED_BY_SERVER, FACILITY_ITF);

    //Create an envelope ID if one has not been specified
  #ifdef CPJNSMTP_MFC_EXTENSIONS 
    if (sENVID.IsEmpty())
      sENVID = CreateNEWENVID();
  
    if (DSNReturnType == CPJNSMTPMessage::HeadersOnly)
      sBuf.Format("MAIL FROM:<%s> RET=HDRS ENVID=%s\r\n", CPJNSMTPAsciiString(pszEmailAddress).operator LPCSTR(), sENVID.operator LPCSTR());
    else
      sBuf.Format("MAIL FROM:<%s> RET=FULL ENVID=%s\r\n", CPJNSMTPAsciiString(pszEmailAddress).operator LPCSTR(), sENVID.operator LPCSTR());
  #else
    if (sENVID.length() == 0)
      sENVID = CreateNEWENVID();

    std::ostringstream ss;
    if (DSNReturnType == CPJNSMTPMessage::HeadersOnly)
    {
    #ifdef _UNICODE
      ss << "MAIL FROM:<" << ATL::CT2A(pszEmailAddress) << "> RET=HDRS ENVID=" << sENVID << "\r\n";
    #else
      ss << "MAIL FROM:<" << pszEmailAddress << "> RET=HDRS ENVID=" << sENVID << "\r\n";
    #endif //#ifdef _UNICODE
    }
    else
    {
    #ifdef _UNICODE
      ss << "MAIL FROM:<" << ATL::CT2A(pszEmailAddress) << "> RET=FULL ENVID=" << sENVID << "\r\n";
    #else
      ss << "MAIL FROM:<" << pszEmailAddress << "> RET=FULL ENVID=" << sENVID << "\r\n";
    #endif //#ifdef _UNICODE
    }
    sBuf = ss.str();
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS 
  }

  return sBuf;  
}

void CPJNSMTPConnection::SendMessage(_Inout_ CPJNSMTPMessage& message)
{
  //paramater validity checking
  ATLASSERT(m_bConnected); //Must be connected to send a message

  //Send the MAIL command
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sBuf(FormMailFromCommand(message.m_From.m_sEmailAddress, message.m_dwDSN, message.m_DSNReturnType, message.m_sENVID));
#else
  CPJNSMTPAsciiString sBuf(FormMailFromCommand(message.m_From.m_sEmailAddress.c_str(), message.m_dwDSN, message.m_DSNReturnType, message.m_sENVID));
#endif //CPJNSMTP_MFC_EXTENSIONS

  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
  #else
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS 

  //check the response to the MAIL command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_MAIL_FROM_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Send the RCPT command, one for each recipient (includes the TO, CC & BCC recipients)

  //We must be sending the message to someone!
#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(message.m_To.GetSize() + message.m_CC.GetSize() + message.m_BCC.GetSize());
#else
  ATLASSERT(message.m_To.size() + message.m_CC.size() + message.m_BCC.size());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //First the "To" recipients
#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<message.m_To.GetSize(); i++)
#else
  for (CPJNSMTPMessage::CAddressArray::size_type i=0; i<message.m_To.size(); i++)
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  {
    CPJNSMTPAddress& recipient = message.m_To[i];
    SendRCPTForRecipient(message.m_dwDSN, recipient);
  }

  //Then the "CC" recipients
#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<message.m_CC.GetSize(); i++)
#else
  for (CPJNSMTPMessage::CAddressArray::size_type i=0; i<message.m_CC.size(); i++)
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  {
    CPJNSMTPAddress& recipient = message.m_CC[i];
    SendRCPTForRecipient(message.m_dwDSN, recipient);
  }

  //Then the "BCC" recipients
#ifdef CPJNSMTP_MFC_EXTENSIONS
  for (INT_PTR i=0; i<message.m_BCC.GetSize(); i++)
#else
  for (CPJNSMTPMessage::CAddressArray::size_type i=0; i<message.m_BCC.size(); i++)
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  {
    CPJNSMTPAddress& recipient = message.m_BCC[i];
    SendRCPTForRecipient(message.m_dwDSN, recipient);
  }

  //Send the DATA command
  static char* szDataCommand = "DATA\r\n";
  try
  {
    _Send(szDataCommand, static_cast<int>(strlen(szDataCommand)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the DATA command
  if (!ReadCommandResponse(354))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_DATA_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Send the Message Header
  CPJNSMTPAsciiString sHeader(message.GetHeader());
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sHeader.operator LPCSTR(), sHeader.GetLength());
  #else
    _Send(sHeader.c_str(), static_cast<int>(sHeader.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //Send the Header / body Separator
  static char* szBodyHeader = "\r\n";
  try
  {
    _Send(szBodyHeader, static_cast<int>(strlen(szBodyHeader)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //Now send the contents of the mail    
  if (message.m_RootPart.GetNumberOfChildBodyParts() || message.m_bMime)
  {
    //Send the root body part (and all its children)
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    SendBodyPart(&message.m_RootPart, TRUE, message.GetCharset());
  #else
    SendBodyPart(&message.m_RootPart, TRUE, message.GetCharset().c_str());
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
  else
  {
    //Do the UTF8 conversion if necessary
    CPJNSMTPAsciiString sBody;
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    if (message.m_RootPart.GetCharset().CompareNoCase(_T("UTF-8")) == 0)
      sBody = CPJNSMTPBodyPart::ConvertToUTF8(message.m_RootPart.GetText());
    else
      sBody = message.m_RootPart.GetText();
  #else
    if (_tcsicmp(message.m_RootPart.GetCharset().c_str(), _T("UTF-8")) == 0)
      sBody = CPJNSMTPBodyPart::ConvertToUTF8(message.m_RootPart.GetText().c_str());
    else
    {
    #ifdef _UNICODE
      sBody = ATL::CT2A(message.m_RootPart.GetText().c_str());
    #else
      sBody = message.m_RootPart.GetText();
    #endif //#ifdef _UNICODE
    }
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  
    CPJNSMTPBodyPart::FixSingleDot(sBody);

    //Send the body
    try
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      _Send(sBody.operator LPCSTR(), sBody.GetLength());
    #else
      _Send(sBody.c_str(), static_cast<int>(sBody.length()));
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS
  }

  //Send the end of message indicator
  static char* szEOM = "\r\n.\r\n";
  try
  {
    _Send(szEOM, static_cast<int>(strlen(szEOM)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the End of Message command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_END_OF_MESSAGE_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

BOOL CPJNSMTPConnection::OnSendProgress(_In_ DWORD /*dwCurrentBytes*/, _In_ DWORD /*dwTotalBytes*/)
{
  //By default just return TRUE to allow the mail to continue to be sent
  return TRUE; 
}

void CPJNSMTPConnection::SendMessage(_In_reads_bytes_opt_(dwMessageSize) const BYTE* pMessage, _In_ DWORD dwMessageSize, _Inout_ CPJNSMTPMessage::CAddressArray& Recipients, _In_ const CPJNSMTPAddress& From, _Inout_ CPJNSMTPAsciiString& sENVID, _In_ DWORD dwSendBufferSize, _In_ DWORD DSN, _In_ CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType)
{
  //paramater validity checking
  ATLASSERT(m_bConnected); //Must be connected to send a message

  //Send the MAIL command
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    CPJNSMTPAsciiString sBuf(FormMailFromCommand(From.m_sEmailAddress, DSN, DSNReturnType, sENVID));
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
  #else
    CPJNSMTPAsciiString sBuf(FormMailFromCommand(From.m_sEmailAddress.c_str(), DSN, DSNReturnType, sENVID));
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the MAIL command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_MAIL_FROM_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Must be sending to someone
#ifdef CPJNSMTP_MFC_EXTENSIONS
  INT_PTR nRecipients = Recipients.GetSize();
#else
  INT_PTR nRecipients = Recipients.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(nRecipients);

  //Send the RCPT command, one for each recipient
  for (INT_PTR i=0; i<nRecipients; i++)
  {
    CPJNSMTPAddress& recipient = Recipients[i];
    SendRCPTForRecipient(DSN, recipient);
  }

  //Send the DATA command
  static char* szDataCommand = "DATA\r\n";
  try
  {
    _Send(szDataCommand, static_cast<int>(strlen(szDataCommand)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the DATA command
  if (!ReadCommandResponse(354))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_DATA_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Read and send the data a chunk at a time
  BOOL bMore = TRUE;
  DWORD dwBytesSent = 0;
  const BYTE* pSendBuf = pMessage; 
  do
  {
    DWORD dwRead = min(dwSendBufferSize, dwMessageSize - dwBytesSent);
    dwBytesSent += dwRead;

    //Call the progress virtual method
    if (OnSendProgress(dwBytesSent, dwMessageSize))
    {
      try
      {
        _Send(pSendBuf, dwRead);
      }
    #ifdef CWSOCKET_MFC_EXTENSIONS
      catch(CWSocketException* pEx)
      {
        DWORD dwError = pEx->m_nError;
        pEx->Delete();
        ThrowPJNSMTPException(dwError, FACILITY_WIN32);
      }
    #else
      catch(CWSocketException& e)
      {
        ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
      }
    #endif //#ifdef CWSOCKET_MFC_EXTENSIONS
    }
    else
    {
      //Throw an exception to indicate that the sending was cancelled
      ThrowPJNSMTPException(ERROR_CANCELLED, FACILITY_WIN32);
    }

    //Prepare for the next time around
    pSendBuf += dwRead;
    bMore = (dwBytesSent < dwMessageSize);
  }
  while (bMore);

  //Send the end of message indicator
  static char* szEOM = "\r\n.\r\n";
  try
  {
    _Send(szEOM, static_cast<int>(strlen(szEOM)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the End of Message command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_END_OF_MESSAGE_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

void CPJNSMTPConnection::SendMessage(_In_z_ LPCTSTR pszMessageOnFile, _Inout_ CPJNSMTPMessage::CAddressArray& Recipients, _In_ const CPJNSMTPAddress& From, _Inout_ CPJNSMTPAsciiString& sENVID, _In_ DWORD dwSendBufferSize, _In_ DWORD DSN, _In_ CPJNSMTPMessage::DSN_RETURN_TYPE DSNReturnType)
{
  //paramater validity checking
  ATLASSERT(m_bConnected); //Must be connected to send a message

  //Open up the file we want to send
  ATL::CAtlFile file;
  HRESULT hr = file.Create(pszMessageOnFile, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
  if (FAILED(hr))
    ThrowPJNSMTPException(hr);

  //Get the length of the file (we only support sending files less than 4GB!)
  ULONGLONG nFileSize = 0;
  hr = file.GetSize(nFileSize);
  if (FAILED(hr))
    CPJNSMTPConnection::ThrowPJNSMTPException(hr);
  if (nFileSize >= ULONG_MAX)
    CPJNSMTPConnection::ThrowPJNSMTPException(IDS_PJNSMTP_FILE_SIZE_TO_SEND_TOO_LARGE, FACILITY_ITF);
  if (nFileSize == 0)
    ThrowPJNSMTPException(IDS_PJNSMTP_CANNOT_SEND_ZERO_BYTE_MESSAGE, FACILITY_ITF);

  //Send the MAIL command
#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(From.m_sEmailAddress.GetLength());
  CPJNSMTPAsciiString sBuf(FormMailFromCommand(From.m_sEmailAddress, DSN, DSNReturnType, sENVID));
  try
  {
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
#else
  ATLASSERT(From.m_sEmailAddress.length());
  CPJNSMTPAsciiString sBuf(FormMailFromCommand(From.m_sEmailAddress.c_str(), DSN, DSNReturnType, sENVID));
  try
  {
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the MAIL command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_MAIL_FROM_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Must be sending to someone
#ifdef CPJNSMTP_MFC_EXTENSIONS
  INT_PTR nRecipients = Recipients.GetSize();
#else
  INT_PTR nRecipients = Recipients.size();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(nRecipients);

  //Send the RCPT command, one for each recipient
  for (INT_PTR i=0; i<nRecipients; i++)
  {
    CPJNSMTPAddress& recipient = Recipients[i];
    SendRCPTForRecipient(DSN, recipient);
  }

  //Send the DATA command
  static char* szDataCommand = "DATA\r\n";
  try
  {
    _Send(szDataCommand, static_cast<int>(strlen(szDataCommand)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the DATA command
  if (!ReadCommandResponse(354))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_DATA_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Allocate a buffer we will use in the sending
  ATL::CHeapPtr<BYTE> sendBuf;
  if (!sendBuf.Allocate(dwSendBufferSize))
    ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32); 

  //Read and send the data a chunk at a time
  BOOL bMore = TRUE;
  DWORD dwBytesSent = 0; 
  do
  {
    //Read the chunk from file
    DWORD dwRead = 0;
    hr = file.Read(sendBuf, dwSendBufferSize, dwRead);
    if (FAILED(hr))
      ThrowPJNSMTPException(hr);
    bMore = (dwRead == dwSendBufferSize);

    //Send the chunk
    if (dwRead)
    {
      dwBytesSent += dwRead;

      //Call the progress virtual method
      if (OnSendProgress(dwBytesSent, static_cast<DWORD>(nFileSize)))
      {
        try
        {
          _Send(sendBuf, dwRead);
        }
      #ifdef CWSOCKET_MFC_EXTENSIONS
        catch(CWSocketException* pEx)
        {
          DWORD dwError = pEx->m_nError;
          pEx->Delete();
          ThrowPJNSMTPException(dwError, FACILITY_WIN32);
        }
      #else
        catch(CWSocketException& e)
        {
          ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
        }
      #endif //#ifdef CWSOCKET_MFC_EXTENSIONS
      }
      else
      {
        //Throw an exception to indicate that the sending was cancelled
        ThrowPJNSMTPException(ERROR_CANCELLED, FACILITY_WIN32);
      }
    }
  }
  while (bMore);

  //Send the end of message indicator
  static char* szEOM = "\r\n.\r\n";
  try
  {
    _Send(szEOM, static_cast<int>(strlen(szEOM)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the End of Message command
  if (!ReadCommandResponse(250))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_END_OF_MESSAGE_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

void CPJNSMTPConnection::SendRCPTForRecipient(_In_ DWORD dwDSN, _Inout_ CPJNSMTPAddress& recipient)
{
  //Validate our parameters
#ifdef CPJNSMTP_MFC_EXTENSIONS
  ATLASSERT(recipient.m_sEmailAddress.GetLength()); //must have an email address for this recipient
#else
  ATLASSERT(recipient.m_sEmailAddress.length()); //must have an email address for this recipient
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //form the command to send
  CPJNSMTPAsciiString sBuf;
  if (dwDSN == CPJNSMTPMessage::DSN_NOT_SPECIFIED)
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    sBuf.Format("RCPT TO:<%s>\r\n", CPJNSMTPAsciiString(recipient.m_sEmailAddress).operator LPCSTR());
  #else
    std::ostringstream ss;
  #ifdef _UNICODE
    ss << "RCPT TO:<" << ATL::CT2A(recipient.m_sEmailAddress.c_str()) << ">\r\n";
  #else
    ss << "RCPT TO:<" << recipient.m_sEmailAddress << ">\r\n";
  #endif //#ifdef _UNICODE
    sBuf = ss.str();
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
  else
  {
    CPJNSMTPAsciiString sNotificationTypes;
    if (dwDSN & CPJNSMTPMessage::DSN_SUCCESS)
      sNotificationTypes = "SUCCESS";
    if (dwDSN & CPJNSMTPMessage::DSN_FAILURE)
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      if (sNotificationTypes.GetLength())
    #else
      if (sNotificationTypes.length())
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
        sNotificationTypes += ",";
      sNotificationTypes += "FAILURE";
    }      
    if (dwDSN & CPJNSMTPMessage::DSN_DELAY)
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      if (sNotificationTypes.GetLength())
    #else
      if (sNotificationTypes.length())
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
        sNotificationTypes += ",";
      sNotificationTypes += "DELAY";
    }  
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    if (sNotificationTypes.IsEmpty()) //Not setting DSN_SUCCESS, DSN_FAILURE or DSN_DELAY in m_DSNReturnType implies we do not want DSN's
  #else
    if (sNotificationTypes.length() == 0) //Not setting DSN_SUCCESS, DSN_FAILURE or DSN_DELAY in m_DSNReturnType implies we do not want DSN's
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      sNotificationTypes += "NEVER";
    
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    sBuf.Format("RCPT TO:<%s> NOTIFY=%s\r\n", CPJNSMTPAsciiString(recipient.m_sEmailAddress).operator LPCSTR(), sNotificationTypes.operator LPCSTR());
  #else
    std::ostringstream ss;
  #ifdef _UNICODE
    ss << "RCPT TO:<" << ATL::CT2A(recipient.m_sEmailAddress.c_str()) << "> NOTIFY=" << sNotificationTypes << "\r\n";
  #else
    ss << "RCPT TO:<" << recipient.m_sEmailAddress << "> NOTIFY=" << sNotificationTypes << "\r\n";
  #endif //#ifdef _UNICODE
    sBuf = ss.str();
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }

  //and send it
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
  #else
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the RCPT command (250 or 251) is considered success
  if (!ReadCommandResponse(250, 251))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_RCPT_TO_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

BOOL CPJNSMTPConnection::ReadCommandResponse(_In_ int nExpectedCode)
{
  CPJNSMTPAsciiString sResponse;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  sResponse.Preallocate(1024);
#else
  sResponse.reserve(1024);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  BOOL bSuccess = ReadResponse(sResponse);

  if (bSuccess)
    bSuccess = (m_nLastCommandResponseCode == nExpectedCode);

  return bSuccess;
}

BOOL CPJNSMTPConnection::ReadCommandResponse(_In_ int nExpectedCode1, _In_ int nExpectedCode2)
{
  CPJNSMTPAsciiString sResponse;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  sResponse.Preallocate(1024);
#else
  sResponse.reserve(1024);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  BOOL bSuccess = ReadResponse(sResponse);

  if (bSuccess)
    bSuccess = (m_nLastCommandResponseCode == nExpectedCode1) || (m_nLastCommandResponseCode == nExpectedCode2);

  return bSuccess;
}

BOOL CPJNSMTPConnection::ReadResponse(_Inout_ CPJNSMTPAsciiString& sResponse)
{
#ifndef CPJNSMTP_NOSSL
  if (GetSocket() != INVALID_SOCKET)
    return ReadSSLResponse(sResponse);
  else
#endif //#ifndef CPJNSMTP_NOSSL
    return ReadPlainTextResponse(sResponse);
}

BOOL CPJNSMTPConnection::ReadPlainTextResponse(_Inout_ CPJNSMTPAsciiString& sResponse)
{
  //must have been created first
  ATLASSERT(m_bConnected);

  static const char* pszTerminator = "\r\n";
  static int nTerminatorLen = 2;

  //retrieve the reponse until we get the full response or a timeout occurs
  BOOL bFoundFullResponse = FALSE;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  int nStartOfLastLine = 0;
  sResponse.Empty();
#else
  CPJNSMTPAsciiString::size_type nStartOfLastLine = 0;
  sResponse.clear();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  while (!bFoundFullResponse)
  {
    //check the socket for readability
    try
    {
      if (!m_Socket.IsReadible(m_dwTimeout)) //A timeout has occured so fail the function call
      {
        //Hive away the last command reponse
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        m_sLastCommandResponse = sResponse;
      #else

      #ifdef _UNICODE
        m_sLastCommandResponse = ATL::CA2T(sResponse.c_str());
      #else
        m_sLastCommandResponse = sResponse;
      #endif //#ifdef _UNICODE
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
        return FALSE;
      }
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS

    //receive the data from the socket
    char sBuf[4096];
    int nData = 0;
    try
    {
      nData = m_Socket.Receive(sBuf, sizeof(sBuf) - 1);
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS

    //Handle a graceful disconnect
    if (nData == 0)
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      //Hive away the last command reponse
      m_sLastCommandResponse = sResponse;
    #else
      //Hive away the last command reponse
    #ifdef _UNICODE
      m_sLastCommandResponse = ATL::CA2T(sResponse.c_str());
    #else
      m_sLastCommandResponse = sResponse;
    #endif //#ifdef _UNICODE
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      ThrowPJNSMTPException(IDS_PJNSMTP_GRACEFUL_DISCONNECT, FACILITY_ITF, m_sLastCommandResponse);
    }

    //NULL terminate the data received
    sBuf[nData] = '\0';

    //Grow the response data
    sResponse += sBuf;
    
    //Check to see if we got a full response yet
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    int nReceived = sResponse.GetLength();
    LPCSTR szResponse = sResponse;
  #else
    CPJNSMTPAsciiString::size_type nReceived = sResponse.length();
    LPCSTR szResponse = sResponse.c_str();
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    BOOL bFoundTerminator = (strncmp(&szResponse[nReceived - nTerminatorLen], pszTerminator, nTerminatorLen) == 0);
    if (bFoundTerminator && (nReceived > 5))
    {
      //Find the start of the last line we have received
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      int i = nReceived - 3;
    #else
      CPJNSMTPAsciiString::size_type i = nReceived - 3;
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      nStartOfLastLine = 0;
      while ((nStartOfLastLine == 0) && (i > 0))
      {
        if ((szResponse[i] == '\r') && (szResponse[i + 1] == '\n'))
          nStartOfLastLine = i + 2;
        else
        {
          //Prepare for the next loop around
          --i;
        }
      }
      bFoundFullResponse = (szResponse[nStartOfLastLine + 3] == ' ');
    }		
  }

#ifdef CPJNSMTP_MFC_EXTENSIONS
  //determine the Numeric response code
  CPJNSMTPAsciiString sCode(sResponse.Mid(nStartOfLastLine, 3));
  sscanf_s(sCode, "%d", &m_nLastCommandResponseCode);

  //Hive away the last command reponse
  m_sLastCommandResponse = sResponse;
#else
  //determine the Numeric response code
  CPJNSMTPAsciiString sCode(sResponse.substr(nStartOfLastLine, 3));
  sscanf_s(sCode.c_str(), "%d", &m_nLastCommandResponseCode);

  //Hive away the last command reponse
#ifdef _UNICODE
  m_sLastCommandResponse = ATL::CA2T(sResponse.c_str());
#else
  m_sLastCommandResponse = sResponse;
#endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  return TRUE;
}

#ifndef CPJNSMTP_NOSSL
BOOL CPJNSMTPConnection::ReadSSLResponse(_Inout_ CPJNSMTPAsciiString& sResponse)
{
  //must have been created first
  ATLASSERT(m_bConnected);

  static const char* pszTerminator = "\r\n";
  static int nTerminatorLen = 2;

#ifdef CPJNSMTP_MFC_EXTENSIONS
  sResponse.Empty();
#else
  sResponse.clear();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //retrieve the reponse until we 
  //get the full response or a timeout occurs
  BOOL bFoundFullResponse = FALSE;
  int nStartOfLastLine = 0;
  while (!bFoundFullResponse)
  {
    //receive the data from the socket
    SSLWrappers::CMessage message;
    SECURITY_STATUS ss = GetEncryptedMessage(message);
    if (ss != SEC_E_OK)
      ThrowPJNSMTPException(ss);

    //Handle a graceful disconnect
    if (message.m_lSize == 0)
    {
    #ifdef CPJNSMTP_MFC_EXTENSIONS
      //Hive away the last command reponse
      m_sLastCommandResponse = sResponse;
    #else
    #ifdef _UNICODE
      m_sLastCommandResponse = ATL::CA2T(sResponse.c_str());
    #else
      m_sLastCommandResponse = sResponse;
    #endif //#ifdef _UNICODE
    #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

      ThrowPJNSMTPException(IDS_PJNSMTP_GRACEFUL_DISCONNECT, FACILITY_ITF, m_sLastCommandResponse);
    }

  #ifdef CPJNSMTP_MFC_EXTENSIONS
    //Append the data to the response
    sResponse.Append(reinterpret_cast<const char*>(message.m_pbyData), message.m_lSize);

    //Check to see if we got a full response yet
    int nReceived = sResponse.GetLength();
    LPCSTR szResponse = sResponse;
  #else
    //Append the data to the response
    sResponse.append(reinterpret_cast<const char*>(message.m_pbyData), reinterpret_cast<const char*>(message.m_pbyData) + message.m_lSize);

    //Check to see if we got a full response yet
    int nReceived = static_cast<int>(sResponse.length());
    LPCSTR szResponse = sResponse.c_str();
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    BOOL bFoundTerminator = (strncmp(&szResponse[nReceived - nTerminatorLen], pszTerminator, nTerminatorLen) == 0);
    if (bFoundTerminator && (nReceived > 5))
    {
      //Find the start of the last line we have received
      int i = nReceived - 3;
      nStartOfLastLine = 0;
      while ((nStartOfLastLine == 0) && (i > 0))
      {
        if ((szResponse[i] == '\r') && (szResponse[i + 1] == '\n'))
          nStartOfLastLine = i + 2;
        else
        {
          //Prepare for the next loop around
          --i;
        }
      }
      bFoundFullResponse = (szResponse[nStartOfLastLine + 3] == ' ');
    }		
  }

#ifdef CPJNSMTP_MFC_EXTENSIONS
  //determine the numeric response code
  CPJNSMTPAsciiString sCode(sResponse.Mid(nStartOfLastLine, 3));
  sscanf_s(sCode, "%d", &m_nLastCommandResponseCode);

  //Hive away the last command reponse
  m_sLastCommandResponse = sResponse;
#else
  //determine the numeric response code
  CPJNSMTPAsciiString sCode(sResponse.substr(nStartOfLastLine, 3));
  sscanf_s(sCode.c_str(), "%d", &m_nLastCommandResponseCode);

  //Hive away the last command reponse
#ifdef _UNICODE
  m_sLastCommandResponse = ATL::CA2T(sResponse.c_str());
#else
  m_sLastCommandResponse = sResponse;
#endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  return TRUE;
}
#endif //#ifndef CPJNSMTP_NOSSL

void CPJNSMTPConnection::AuthLogin(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszPassword)
{
  //Validate our parameters
  ATLASSERT(pszUsername != NULL);
  ATLASSERT(pszPassword != NULL);

  CPJNSMTPAsciiString sAsciiLastCommandString;
  CPJNSMPTBase64Encode encode;

  //Send the AUTH LOGIN command if required
  if (!m_bDoInitialClientResponse)
  {
    static char* szAuth = "AUTH LOGIN\r\n";
    try
    {
      _Send(szAuth, static_cast<int>(strlen(szAuth)));
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS

    //check the response to the AUTH LOGIN command
    if (!ReadCommandResponse(334))
      ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_LOGIN_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

    //Check that the response has a username request in it
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    sAsciiLastCommandString = m_sLastCommandResponse.Mid(4);
  #else
  #ifdef _UNICODE
    sAsciiLastCommandString = ATL::CT2A(m_sLastCommandResponse.substr(4).c_str());
  #else
    sAsciiLastCommandString = m_sLastCommandResponse.substr(4);
  #endif //#ifdef _UNICODE
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    encode.Decode(sAsciiLastCommandString);
  #else
    encode.Decode(sAsciiLastCommandString.c_str());
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    if (_strnicmp(encode.Result(), "username:", 9) != 0)
      ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_LOGIN_USERNAME_REQUEST, FACILITY_ITF, m_sLastCommandResponse);
  }

  //send base64 encoded username
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiUsername(pszUsername);
  encode.Encode(sAsciiUsername, ATL_BASE64_FLAG_NOCRLF);
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiUsername(ATL::CT2A(pszUsername).operator LPSTR());
#else
  CPJNSMTPAsciiString sAsciiUsername(pszUsername);
#endif //#ifdef _UNICODE
  encode.Encode(sAsciiUsername.c_str(), ATL_BASE64_FLAG_NOCRLF);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sUser;
  if (m_bDoInitialClientResponse)
    sUser += "AUTH LOGIN ";
  sUser += encode.Result();
  sUser += "\r\n";
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sUser.operator LPCSTR(), sUser.GetLength());
  #else
    _Send(sUser.c_str(), static_cast<int>(sUser.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check the response to the username 
  if (!ReadCommandResponse(334))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_LOGIN_USERNAME_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Check that the response has a password request in it
#ifdef CPJNSMTP_MFC_EXTENSIONS
  sAsciiLastCommandString = m_sLastCommandResponse.Mid(4);
  encode.Decode(sAsciiLastCommandString);
#else
#ifdef _UNICODE
  sAsciiLastCommandString = ATL::CT2A(m_sLastCommandResponse.substr(4).c_str());
#else
  sAsciiLastCommandString = m_sLastCommandResponse.substr(4);
#endif //#ifdef _UNICODE
  encode.Decode(sAsciiLastCommandString.c_str());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  if (_strnicmp(encode.Result(), "password:", 9) != 0)
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_LOGIN_PASSWORD_REQUEST, FACILITY_ITF, m_sLastCommandResponse);

  //send password as base64 encoded
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiPassword(pszPassword);
  encode.Encode(sAsciiPassword, ATL_BASE64_FLAG_NOCRLF);
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiPassword(ATL::CT2A(pszPassword).operator LPSTR());
#else
  CPJNSMTPAsciiString sAsciiPassword(pszPassword);
#endif //#ifdef _UNICODE
  encode.Encode(sAsciiPassword.c_str(), ATL_BASE64_FLAG_NOCRLF);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sPwd(encode.Result());
  sPwd += "\r\n";
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sPwd.operator LPCSTR(), sPwd.GetLength());
  #else
    _Send(sPwd.c_str(), static_cast<int>(sPwd.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check if authentication is successful
  if (!ReadCommandResponse(235))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_LOGIN_PASSWORD_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

void CPJNSMTPConnection::AuthPlain(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszPassword)
{
  //Validate our parameters
  ATLASSERT(pszUsername != NULL);
  ATLASSERT(pszPassword != NULL);

  //Send the AUTH PLAIN command if required
  if (!m_bDoInitialClientResponse)
  {
    static char* szAuth = "AUTH PLAIN\r\n";
    try
    {
      _Send(szAuth, static_cast<int>(strlen(szAuth)));
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS

    //Check the response to the AUTH PLAIN command
    if (!ReadCommandResponse(334))
      ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_PLAIN_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
  }

  //Send the username and password in the base64 encoded AUTH PLAIN format of "authid\0userid\0passwd".
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiUserName(pszUsername);
  CPJNSMTPAsciiString sAsciiPassword(pszPassword);
  int nAsciiUserNameLen = sAsciiUserName.GetLength();
  int nAsciiPasswordLen = sAsciiPassword.GetLength();
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiUserName(ATL::CT2A(pszUsername).operator LPSTR());
  CPJNSMTPAsciiString sAsciiPassword(ATL::CT2A(pszPassword).operator LPSTR());
#else
  CPJNSMTPAsciiString sAsciiUserName(pszUsername);
  CPJNSMTPAsciiString sAsciiPassword(pszPassword);
#endif //#ifdef _UNICODE
  int nAsciiUserNameLen = static_cast<int>(sAsciiUserName.length());
  int nAsciiPasswordLen = static_cast<int>(sAsciiPassword.length());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  int nAuthLen = (nAsciiUserNameLen * 2) + nAsciiPasswordLen + 2;

  //Allocate some heap space for the auth request packet
  ATL::CHeapPtr<char> authRequest;
  if (!authRequest.Allocate(nAuthLen))
    ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32); 
  
  //Form the auth request
#ifdef CPJNSMTP_MFC_EXTENSIONS
  strcpy_s(authRequest.m_pData, nAuthLen, sAsciiUserName);
  strcpy_s(authRequest.m_pData + nAsciiUserNameLen + 1, nAuthLen - nAsciiUserNameLen - 1, sAsciiUserName);
  memcpy_s(authRequest.m_pData + (2* nAsciiUserNameLen) + 2, nAuthLen - (2 * nAsciiUserNameLen) - 2, sAsciiPassword, nAsciiPasswordLen);
#else
  strcpy_s(authRequest.m_pData, nAuthLen, sAsciiUserName.c_str());
  strcpy_s(authRequest.m_pData + nAsciiUserNameLen + 1, nAuthLen - nAsciiUserNameLen - 1, sAsciiUserName.c_str());
  memcpy_s(authRequest.m_pData + (2*nAsciiUserNameLen) + 2, nAuthLen - (2*nAsciiUserNameLen) - 2, sAsciiPassword.c_str(), nAsciiPasswordLen);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMPTBase64Encode encode;
  encode.Encode(reinterpret_cast<const BYTE*>(authRequest.m_pData), nAuthLen, ATL_BASE64_FLAG_NOCRLF);
  SecureZeroMemory(authRequest.m_pData, nAuthLen);
  
  CPJNSMTPAsciiString sAuth;
  if (m_bDoInitialClientResponse)
    sAuth += "AUTH PLAIN ";
  sAuth += encode.Result();
  sAuth += "\r\n";
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sAuth.operator LPCSTR(), sAuth.GetLength());
  #else
    _Send(sAuth.c_str(), static_cast<int>(sAuth.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check if authentication is successful
  if (!ReadCommandResponse(235))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_PLAIN_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

void CPJNSMTPConnection::AuthCramMD5(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszPassword)
{
  //Validate our parameters
  ATLASSERT(pszUsername != NULL);
  ATLASSERT(pszPassword != NULL);

  //Send the AUTH CRAM-MD5 command
  static char* szAuth = "AUTH CRAM-MD5\r\n";
  try
  {
    _Send(szAuth, static_cast<int>(strlen(szAuth)));
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //Check the response to the AUTH CRAM-MD5 command
  if (!ReadCommandResponse(334))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_CRAM_MD5_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Get the base64 decoded challange 
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiLastCommandString(m_sLastCommandResponse.Mid(4));
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiLastCommandString(ATL::CT2A(m_sLastCommandResponse.substr(4).c_str()));
#else
  CPJNSMTPAsciiString sAsciiLastCommandString(m_sLastCommandResponse.substr(4));
#endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMPTBase64Encode encode;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  encode.Decode(sAsciiLastCommandString);
#else
  encode.Decode(sAsciiLastCommandString.c_str());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  LPCSTR pszChallenge = encode.Result();

  //generate the MD5 digest from the challenge and password
  CPJNMD5 hmac;
  CPJNMD5Hash hash;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiPassword(pszPassword);
  if (!hmac.HMAC(reinterpret_cast<const BYTE*>(pszChallenge), static_cast<DWORD>(strlen(pszChallenge)), reinterpret_cast<const BYTE*>(sAsciiPassword.operator LPCSTR()), sAsciiPassword.GetLength(), hash))
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiPassword(ATL::CT2A(pszPassword).operator LPSTR());
#else
  CPJNSMTPAsciiString sAsciiPassword(pszPassword);
#endif //#ifdef _UNICODE
  if (!hmac.HMAC(reinterpret_cast<const BYTE*>(pszChallenge), static_cast<DWORD>(strlen(pszChallenge)), reinterpret_cast<const BYTE*>(sAsciiPassword.c_str()), static_cast<DWORD>(sAsciiPassword.length()), hash))
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
    ThrowPJNSMTPException(GetLastError(), FACILITY_WIN32);
  
  //make the CRAM-MD5 response
  CPJNSMTPString sCramDigest(pszUsername);
  sCramDigest += _T(" ");
#ifdef CPJNMD5_MFC_EXTENSIONS
  sCramDigest += hash.Format(FALSE); //CRAM-MD5 requires a lowercase hash
#else
  sCramDigest += hash.Format(FALSE).c_str(); //CRAM-MD5 requires a lowercase hash
#endif //#ifdef CPJNMD5_MFC_EXTENSIONS
    
  //send the digest response
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiCramDigest(sCramDigest);
  encode.Encode(sAsciiCramDigest, ATL_BASE64_FLAG_NOCRLF);
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiCramDigest(ATL::CT2A(sCramDigest.c_str()));
#else
  CPJNSMTPAsciiString sAsciiCramDigest(sCramDigest);
#endif //#ifdef _UNICODE
  encode.Encode(sAsciiCramDigest.c_str(), ATL_BASE64_FLAG_NOCRLF);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sEncodedDigest(encode.Result());
  sEncodedDigest += "\r\n";
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sEncodedDigest.operator LPCSTR(), sEncodedDigest.GetLength());
  #else
    _Send(sEncodedDigest.c_str(), static_cast<int>(sEncodedDigest.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check if authentication is successful
  if (!ReadCommandResponse(235))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_CRAM_MD5_DIGEST_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
}

void CPJNSMTPConnection::AuthXOAUTH2(_In_z_ LPCTSTR pszUsername, _In_z_ LPCTSTR pszAccessToken)
{
  //Validate our parameters
  ATLASSERT(pszUsername != NULL);
  ATLASSERT(pszAccessToken != NULL);

  //Send the AUTH XOAUTH2 command if required
  if (!m_bDoInitialClientResponse)
  {
    static char* szAuth = "AUTH XOAUTH2\r\n";
    try
    {
      _Send(szAuth, static_cast<int>(strlen(szAuth)));
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch (CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch (CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS

    //Check the response to the AUTH XOAUTH2 command
    if (!ReadCommandResponse(334))
      ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_XOAUTH2_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
  }

  //Send the username and access token in the base64 encoded XOAUTH2 format
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiUserName(pszUsername);
  CPJNSMTPAsciiString sAsciiAccessToken(pszAccessToken);
  int nAsciiUserNameLen = sAsciiUserName.GetLength();
  int nAsciiAccessTokenLen = sAsciiAccessToken.GetLength();
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiUserName(ATL::CT2A(pszUsername).operator LPSTR());
  CPJNSMTPAsciiString sAsciiAccessToken(ATL::CT2A(pszAccessToken).operator LPSTR());
#else
  CPJNSMTPAsciiString sAsciiUserName(pszUsername);
  CPJNSMTPAsciiString sAsciiAccessToken(pszAccessToken);
#endif //#ifdef _UNICODE
  int nAsciiUserNameLen = static_cast<int>(sAsciiUserName.length());
  int nAsciiAccessTokenLen = static_cast<int>(sAsciiAccessToken.length());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  int nAuthLen = nAsciiUserNameLen + nAsciiAccessTokenLen + 20;
  
  //Allocate some heap space for the auth request packet
  ATL::CHeapPtr<char> authRequest;
  if (!authRequest.Allocate(nAuthLen))
    ThrowPJNSMTPException(ERROR_OUTOFMEMORY, FACILITY_WIN32); 

  //Form the auth request
  memcpy_s(authRequest.m_pData, nAuthLen, "user=", 5);
#ifdef CPJNSMTP_MFC_EXTENSIONS
  memcpy_s(authRequest.m_pData + 5, nAuthLen - 5, sAsciiUserName, nAsciiUserNameLen);
#else
  memcpy_s(authRequest.m_pData + 5, nAuthLen - 5, sAsciiUserName.c_str(), nAsciiUserNameLen);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  authRequest.m_pData[5 + nAsciiUserNameLen] = 0x01;
  memcpy_s(authRequest.m_pData + 6 + nAsciiUserNameLen, nAuthLen - 6 - nAsciiUserNameLen, "auth=Bearer ", 12);
#ifdef CPJNSMTP_MFC_EXTENSIONS
  memcpy_s(authRequest.m_pData + 18 + nAsciiUserNameLen, nAuthLen - 18 - nAsciiUserNameLen, sAsciiAccessToken, nAsciiAccessTokenLen);
#else
  memcpy_s(authRequest.m_pData + 18 + nAsciiUserNameLen, nAuthLen - 18 - nAsciiUserNameLen, sAsciiAccessToken.c_str(), nAsciiAccessTokenLen);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  authRequest.m_pData[18 + nAsciiUserNameLen + nAsciiAccessTokenLen] = 0x01;
  authRequest.m_pData[19 + nAsciiUserNameLen + nAsciiAccessTokenLen] = 0x01;

  CPJNSMPTBase64Encode encode;
  encode.Encode(reinterpret_cast<const BYTE*>(authRequest.m_pData), nAuthLen, ATL_BASE64_FLAG_NOCRLF);
  SecureZeroMemory(authRequest.m_pData, nAuthLen);
  
  CPJNSMTPAsciiString sAuth;
  if (m_bDoInitialClientResponse)
    sAuth += "AUTH XOAUTH2 ";
  sAuth += encode.Result();
  sAuth += "\r\n";
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sAuth.operator LPCSTR(), sAuth.GetLength());
  #else
    _Send(sAuth.c_str(), static_cast<int>(sAuth.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check if authentication is successful
  if (!ReadCommandResponse(235))
  {
    try
    {
      _Send("\r\n", 2);
    }
  #ifdef CWSOCKET_MFC_EXTENSIONS
    catch(CWSocketException* pEx)
    {
      DWORD dwError = pEx->m_nError;
      pEx->Delete();
      ThrowPJNSMTPException(dwError, FACILITY_WIN32);
    }
  #else
    catch(CWSocketException& e)
    {
      ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
    }
  #endif //#ifdef CWSOCKET_MFC_EXTENSIONS
    ReadCommandResponse(-1);
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_XOAUTH2_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);
  }
}

CPJNSMTPConnection::ConnectToInternetResult CPJNSMTPConnection::ConnectToInternet()
{
  //Check to see if an internet connection already exists.
  //bInternet = TRUE  internet connection exists.
  //bInternet = FALSE internet connection does not exist
  DWORD dwFlags = 0;
  BOOL bInternet = InternetGetConnectedState(&dwFlags, 0);
  if (!bInternet)
  {
    //Attempt to establish internet connection, probably
    //using Dial-up connection. CloseInternetConnection() should be called when
    //as some time to drop the dial-up connection.
    DWORD dwResult = InternetAttemptConnect(0);
    if (dwResult != ERROR_SUCCESS)
    {
      SetLastError(dwResult);
      return CTIR_Failure;
    }
    else
      return CTIR_NewConnection;
  }
  else
    return CTIR_ExistingConnection;
}

BOOL CPJNSMTPConnection::CloseInternetConnection()
{
  //Make sure any connection through a modem is 'closed'.
  return InternetAutodialHangup(0);
}

#ifndef CPJNSMTP_NONTLM
SECURITY_STATUS CPJNSMTPConnection::NTLMAuthPhase1(_In_reads_bytes_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf)
{
  //Send the AUTH NTLM command with the initial data
  CPJNSMPTBase64Encode encode;
  encode.Encode(pBuf, cbBuf, ATL_BASE64_FLAG_NOCRLF);
  
  CPJNSMTPAsciiString sBuf("AUTH NTLM ");
  sBuf += encode.Result();
  sBuf += "\r\n";
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
  #else
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  return SEC_E_OK;
}

SECURITY_STATUS CPJNSMTPConnection::NTLMAuthPhase2(_Inout_updates_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf, _Out_ DWORD* pcbRead)
{
  //check the response to the AUTH NTLM command
  if (!ReadCommandResponse(334))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_NTLM_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  //Decode the last response
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sAsciiLastCommandString(m_sLastCommandResponse.Mid(4));
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sAsciiLastCommandString(ATL::CT2A(m_sLastCommandResponse.substr(4).c_str()));
#else
  CPJNSMTPAsciiString sAsciiLastCommandString(m_sLastCommandResponse.substr(4));
#endif //#ifdef _UNICODE
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMPTBase64Encode encode;
#ifdef CPJNSMTP_MFC_EXTENSIONS
  encode.Decode(sAsciiLastCommandString);
#else
  encode.Decode(sAsciiLastCommandString.c_str());
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //Store the results in the output parameters
  *pcbRead = encode.ResultSize();
  if (*pcbRead >= cbBuf)
    return SEC_E_INSUFFICIENT_MEMORY;
  memcpy_s(pBuf, cbBuf, encode.Result(), *pcbRead);

  return SEC_E_OK;
}

SECURITY_STATUS CPJNSMTPConnection::NTLMAuthPhase3(_In_reads_bytes_(cbBuf) PBYTE pBuf, _In_ DWORD cbBuf)
{
  //send base64 encoded version of the data
  CPJNSMPTBase64Encode encode;
  encode.Encode(pBuf, cbBuf, ATL_BASE64_FLAG_NOCRLF);

  CPJNSMTPAsciiString sBuf(encode.Result());
  sBuf += "\r\n";
  try
  {
  #ifdef CPJNSMTP_MFC_EXTENSIONS
    _Send(sBuf.operator LPCSTR(), sBuf.GetLength());
  #else
    _Send(sBuf.c_str(), static_cast<int>(sBuf.length()));
  #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
  }
#ifdef CWSOCKET_MFC_EXTENSIONS
  catch(CWSocketException* pEx)
  {
    DWORD dwError = pEx->m_nError;
    pEx->Delete();
    ThrowPJNSMTPException(dwError, FACILITY_WIN32);
  }
#else
  catch(CWSocketException& e)
  {
    ThrowPJNSMTPException(e.m_nError, FACILITY_WIN32);
  }
#endif //#ifdef CWSOCKET_MFC_EXTENSIONS

  //check if authentication is successful
  if (!ReadCommandResponse(235))
    ThrowPJNSMTPException(IDS_PJNSMTP_UNEXPECTED_AUTH_NTLM_RESPONSE, FACILITY_ITF, m_sLastCommandResponse);

  return SEC_E_OK;
}
#endif //#ifdef CPJNSMTP_NONTLM

BOOL CPJNSMTPConnection::MXLookup(_In_z_ LPCTSTR lpszHostDomain, _Inout_ CPJNSMTPStringArray& arrHosts, _Inout_ CPJNSMTPWordArray& arrPreferences, _In_ WORD fOptions, _In_opt_ PIP4_ARRAY aipServers)
{
  //What will be the return value
  BOOL bSuccess = FALSE;

  //Reset the output parameter arrays
#ifdef CPJNSMTP_MFC_EXTENSIONS
  arrHosts.RemoveAll();
  arrPreferences.RemoveAll();
#else
  arrHosts.clear();
  arrPreferences.clear();
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  //Do the DNS MX lookup  
  PDNS_RECORD pRec = NULL;
  DNS_STATUS status = DnsQuery(lpszHostDomain, DNS_TYPE_MX, fOptions, aipServers, &pRec, NULL);
  if (status == ERROR_SUCCESS)
  {
    bSuccess = TRUE;

    PDNS_RECORD pRecFirst = pRec;
    while (pRec != NULL)
    {
      //We're only interested in MX records
      if (pRec->wType == DNS_TYPE_MX)
      {
      #ifdef CPJNSMTP_MFC_EXTENSIONS
        arrHosts.Add(pRec->Data.MX.pNameExchange);
        arrPreferences.Add(pRec->Data.MX.wPreference);
      #else
        arrHosts.push_back(pRec->Data.MX.pNameExchange);
        arrPreferences.push_back(pRec->Data.MX.wPreference);
      #endif //#ifdef CPJNSMTP_MFC_EXTENSIONS
      }
      pRec = pRec->pNext;
    }
    if (pRecFirst != NULL)
      DnsRecordListFree(pRecFirst, DnsFreeRecordList);
  }
  else
    SetLastError(status);

  return bSuccess;
}

CPJNSMTPAsciiString CPJNSMTPConnection::CreateNEWENVID()
{
  //Remove all but the hex digits
#ifdef CPJNSMTP_MFC_EXTENSIONS
  CPJNSMTPAsciiString sEnvID(CPJNSMTPBodyPart::CreateGUID());
  sEnvID.Remove('{');
  sEnvID.Remove('}');
  sEnvID.Remove('-');
  sEnvID.MakeUpper();
#else
#ifdef _UNICODE
  CPJNSMTPAsciiString sEnvID(ATL::CT2A(CPJNSMTPBodyPart::CreateGUID().c_str()));
#else
  CPJNSMTPAsciiString sEnvID(CPJNSMTPBodyPart::CreateGUID());
#endif //#ifdef _UNICODE
  sEnvID.erase(std::remove(sEnvID.begin(), sEnvID.end(), '{'), sEnvID.end());
  sEnvID.erase(std::remove(sEnvID.begin(), sEnvID.end(), '}'), sEnvID.end());
  sEnvID.erase(std::remove(sEnvID.begin(), sEnvID.end(), '-'), sEnvID.end());
  std::transform(sEnvID.begin(), sEnvID.end(), sEnvID.begin(), _totupper);
#endif //#ifdef CPJNSMTP_MFC_EXTENSIONS

  return sEnvID;
}
