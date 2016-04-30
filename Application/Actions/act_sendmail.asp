<!--#include virtual="/framework/configuration/email.config"-->
<!--#include virtual="/framework/elements/miscfunctions.framework"-->
<!--#include virtual="/framework/elements/sha1encryption.framework"-->

<%
// ------------------------------------------------------------------------------------------------------------------------------
// ' Define Variables
// ------------------------------------------------------------------------------------------------------------------------------

  Mail_Query          = Request.Form
	Mail_Query          = Replace( Mail_Query, "&", ";" )
	Mail_Query          = Replace( Mail_Query, "=", ":" )
	Mail_Query          = Replace( Mail_Query, ";", vbcrlf )
	Mail_Query          = UrlDecode( Mail_Query )
	
	Mail_Name           = ParseCircuit( "name", Mail_Query )
	Mail_Email          = ParseCircuit( "email", Mail_Query )
	Mail_Company        = ParseCircuit( "company", Mail_Query )
	Mail_Address1       = ParseCircuit( "address1", Mail_Query )
	Mail_Address2       = ParseCircuit( "address2", Mail_Query )
	Mail_Address3       = ParseCircuit( "address3", Mail_Query )
	Mail_Type1          = ParseCircuit( "type1", Mail_Query )
	Mail_Type2          = ParseCircuit( "type2", Mail_Query )
	Mail_Design         = ParseCircuit( "design", Mail_Query )
	Mail_Build          = ParseCircuit( "build", Mail_Query )
	Mail_Manage         = ParseCircuit( "manage", Mail_Query )
	Mail_Date           = ParseCircuit( "date", Mail_Query )
	Mail_Budget         = ParseCircuit( "budget", Mail_Query )
	Mail_Requirements   = ParseCircuit( "requirements", Mail_Query )
	Mail_Message        = ParseCircuit( "message", Mail_Query )
	
	o_Component         = RequiredComponent
	o_MailServer        = Email_Server
	o_MailPort          = Email_Port
	o_MailAuth          = Email_Username
	o_MailPass          = Email_UserPass
	o_MailRCPT          = Email_Recipient
	o_MailName          = Email_From
	o_MailSubject       = "Enquiry from your website."
	o_Queuing           = False
	o_UseSMTPSSL        = Email_UseSMTPSSL
	o_ConnTimeOut       = Email_SMTPTimeOut
	o_SMTPAuth          = Email_SMTPAuth
	o_MessageId         = Sha1( Timer() & Rnd() )
	o_TruncUser         = Left( o_MailAuth, 3 )
	o_TruncPassword     = Left( o_MailPass, 3 )
	o_TruncRCPT         = Right( o_MailRCPT, 15 )
	o_PassLen           = Len( o_MailPass )
	o_UserLen           = Len( o_MailAuth )
	o_RCPTLen           = Len( o_MailRCPT )
	
	For i = 1 to o_PassLen - 3
	  o_PassTrunc = o_PassTrunc & "*"
	Next
	
	For i = 1 to o_UserLen - 3
	  o_UserTrunc = o_UserTrunc & "*"
	Next
	
	For i = 1 to o_RCPTLen - 15
	  o_RCPTTrunc = o_RCPTTrunc & "*"
	Next
	
	o_TruncPassword = o_TruncPassword & o_PassTrunc
	o_TruncUsername = o_TruncUser & o_UserTrunc
	o_TruncRCPT     = o_RCPTTrunc & o_TruncRCPT
	
	If o_Component = 1 Then
	  HandlerComponent = "CDONTS OBJECT"
	Else
	  HandlerComponent = "PERSITS MAILER OBJECT"
	End If
	
// ------------------------------------------------------------------------------------------------------------------------------
// ' Build Message Body
// ------------------------------------------------------------------------------------------------------------------------------

  Mail_Body = "Subject: Message from website" & Chr(13) & _
	            "------------------------------------------------------------------------" & Chr(13) & Chr(13) & _
							" Contact name: "       & Mail_Name       & Chr(13) & _
							" Contact email:"       & Mail_Email      & Chr(13) & _
							" Company name:"        & Mail_Company    & Chr(13) & Chr(13) & _
							" Address 1:"           & Mail_Address1   & Chr(13) & _
							" Address 2:"           & Mail_Address2   & Chr(13) & _
							" Address 3:"           & Mail_Address3   & Chr(13) & Chr(13) & _
							" Existing business:"   & Mail_Type1      & Chr(13) & _
							" Proposed business:"   & Mail_Type2      & Chr(13) & _
							" Interior design:"     & Mail_Design     & Chr(13) & _
							" Design & build: "     & Mail_Build      & Chr(13) & _
							" Project management:"  & Mail_Manage     & Chr(13) & _
							" Completion date:"     & Mail_Date       & Chr(13) & _
							" Budget:"              & Mail_Budget     & Chr(13) & Chr(13) & _
							" Additional Requirements:"               & Chr(13) & _
							"" & Mail_Requirements                    & Chr(13) & Chr(13) & _
							" Quick message:"                         & Chr(13) & _
							"" & Mail_Message                         & Chr(13) & Chr(13) & _
							"------------------------------------------------------------------------"  & Chr(13) & _
							" Server protocol: " & Request.ServerVariables("SERVER_PROTOCOL")           & Chr(13) & _
							" Remote host: " & Request.ServerVariables("REMOTE_HOST")                   & Chr(13) & _
							" Remote IP address:" & Request.ServerVariables("REMOTE_ADDR")              & Chr(13) & _
							" Message ID:" & Sha1( Timer() & Rnd() )                                    & Chr(13) & _
							" TimeStamp:"  & Now & Chr(13) & _
							" Handler Component:" & HandlerComponent & Chr(13) & _
							"------------------------------------------------------------------------"  & Chr(13)
							
							'Response.Write Mail_Body

// ------------------------------------------------------------------------------------------------------------------------------
// ' Validate Form
// ------------------------------------------------------------------------------------------------------------------------------

  If Mail_Name = "" Then
	  Proceed = 0 
	  ErrCode = 1
		ErrText = "<h2>Your name is required but was not provided.<br/>Please click the back button on your browser to fix this error.</h2>"
	ElseIf Mail_Email = "" Then
	  Proceed = 0
		ErrText = "<h2>You entered an invalid email address.<br/>Please click the back button on your browser to fix this error.</h2>"
		ErrCode = 2
	ElseIf Instr( Mail_Email,"@" ) = 0 Then
	  Proceed = 0
		ErrCode = 2
		ErrText = "<h2>You entered an invalid email address.<br/>Please click the back button on your browser to fix this error.</h2>"
	ElseIf Instr( Mail_Email,"." ) = "" Then
	  Proceed = 0
	  ErrCode = 2
		ErrText = "<h2>You entered an invalid email address.<br/>Please click the back button on your browser to fix this error.</h2>"
	ElseIf Mail_Message = "" Then
	  Proceed = 0
		ErrCode = 4
		ErrText = "<h2>The message field cannot be left blank.<br/>Please click the back button on your browser to fix this error.</h2>"
	ElseIf Err.Number <> 0 Then
	  Proceed = 0
		ErrCode = 5
		ErrText = Err.Number & "-" & Err.Description
	Else
	  Proceed = 1
		ErrCode = 0
		
	End If
	
// ------------------------------------------------------------------------------------------------------------------------------
// ' Build form response
// ------------------------------------------------------------------------------------------------------------------------------

  'On Error Resume Next ' Suppress errors
	
	

  If Proceed = 1 AND ERR.Number = 0 Then
	
	  SuccessNotice = "<td width='170' valign='top'>" & _
		                 "<center><img src='/logo1.gif'/></center>" & _
									   "<font face='Times' size='2' color='#C09F7E'>" & _
									   "<p><a href='/index.html'><img border='0' src='/homeimage.gif'/></a></p>" & _
									   "<p><a href='/design.shtml'><img border='0' src='/designservimage.gif'/></a></p>" & _
									   "<p><a href='/galleries.shtml'><img border='0' src='/designgalleriesimage.gif'/></a></p>" & _
									   "<p><a href='/client.html'></a><img border='0' src='/clientrefimage.gif'/></p>" & _
									   "<p><a href='/services.shtml'><img border='0' src='/assocservimage.gif'/></a></p>" & _
									   "</font>" & _
		                 "</td>" & _
										 "<td style='padding-left:30px;'>" & _
										 "<blockquote><font face='times' size='6'><b>Thank You</b></font></blockquote>" & _
										 "<blockquote>" & _
										 "Thank you for your interest. You will be contacted by someone from David Ostick shortly." & _
										 "<br/><a href='/index.html'>Click here</a> to return to the David Ostick homepage." & _
										 "</blockquote>" & _
										 "</td>"
										 
	Else
	
	  SuccessNotice =  "<td width='170' valign='top'>" & _
		                 "<center><img src='/logo1.gif'/></center>" & _
									   "<font face='Times' size='2' color='#C09F7E'>" & _
									   "<p><a href='/index.html'><img border='0' src='/homeimage.gif'/></a></p>" & _
									   "<p><a href='/design.shtml'><img border='0' src='/designservimage.gif'/></a></p>" & _
									   "<p><a href='/galleries.shtml'><img border='0' src='/designgalleriesimage.gif'/></a></p>" & _
									   "<p><a href='/client.html'></a><img border='0' src='/clientrefimage.gif'/></p>" & _
									   "<p><a href='/services.shtml'><img border='0' src='/assocservimage.gif'/></a></p>" & _
									   "</font>" & _
		                 "</td>" & _
										 "<td style='padding-left:30px;'>" & _
										 "<blockquote>" & ErrText & "</blockquote>" & _
										 "</td>"
	
	End If
	
	Form_Header = "<html><head></head>" & _
	              "<body bgcolor='#C09F7E' background='/bg2.gif' text='#652501'>" & _
								"<table><tr>"
								
								
	Form_Footer = "</tr></table></body></html>"

// ------------------------------------------------------------------------------------------------------------------------------
// ' Send mail using CDONTS component
// ------------------------------------------------------------------------------------------------------------------------------
  
	If Proceed = 1 AND o_Component = 1 Then
	
	  Schema = "http://schemas.microsoft.com/cdo/configuration"
		
		'On Error Resume Next ' Suppress Errors
		
		Set MailObject = CreateObject("CDO.Message")
		MailObject.Configuration.Fields.Item( Schema & "/sendusing" )             = 2
		MailObject.Configuration.Fields.Item( Schema & "/smtpserver" )            = o_MailServer
		MailObject.Configuration.Fields.Item( Schema & "/smtpserverport" )        = o_MailPort
		MailObject.Configuration.Fields.Item( Schema & "/smtpusessl" )            = o_UseSMTPSSL
		MailObject.Configuration.Fields.Item( Schema & "/smtpconnectiontimeout" ) = o_ConnTimeOut
		
	  If o_SMTPAuth = 1 Then
	    MailObject.Configuration.Fields.Item( Schema & "/sendusername" )          = o_MailAuth
		  MailObject.Configuration.Fields.Item( Schema & "/sendpassword" )          = o_MailPass
	  End If
		
		o_MailBody = Mail_Body
		
		MailObject.Configuration.Fields.Update
		MailObject.To         = Email_Recipient
		MailObject.Subject    = o_MailSubject
		MailObject.From       = Mail_Email
		MailObject.TextBody   = o_MailBody
		MailObject.ReplyTo    = o_MailAuth
		
		On Error Resume Next '  Suppress Errors
		MailObject.Send
		
		Set MailObject = Nothing
	
	End If


// ------------------------------------------------------------------------------------------------------------------------------
// ' Send mail using Persits component
// ------------------------------------------------------------------------------------------------------------------------------

  If Proceed = 1 AND o_Component = 2 Then
	
	  'On Error Resume Next
		Set MailObject = CreateObject("Persits.MailSender") 
		
		o_MailBody = Mail_Body
		
		MailObject.Host       = o_MailServer
		MailObject.Port       = o_MailPort
		MailObject.Username   = o_MailAuth
		MailObject.Password   = o_MailPass
		MailObject.From       = Email_Username
		MailObject.FromName   = "DavidOstick.co.uk - POSTMASTER"
		MailObject.AddAddress o_MailRCPT
		MailObject.AddReplyTo o_MailName
		MailObject.Subject    = "Enquiry from your website" 
		MailObject.Body       = Mail_Body
		MailObject.IsHtml     = False
		MailObject.Queue      = False
		
		MailObject.Send
		Set MailObject = Nothing
	
	End If

// ------------------------------------------------------------------------------------------------------------------------------
	
	If Err.Number <> 0 Then
	
	Response.Write "<pre>" & Err.Description & " - " & Err.Number & "</pre>"
	
	Else
	
	Response.Write Form_Header & SuccessNotice & Form_Footer

	End If
	
// ------------------------------------------------------------------------------------------------------------------------------
%>