<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>

<!--#INCLUDE file="inc_config.asp"-->
<!--#INCLUDE file="inc_sha256.asp"-->

<%
Function RestrictPageAccesToLevel(ByVal intTempPageLevel)
	if CInt(Session("UserLevel")) < CInt(intTempPageLevel) then
		if Request.Cookies(ajlogin_cookiename)<>"" then
			strCookieUserName = Request.Cookies(ajlogin_cookiename)("strUserName")
			strCookiePassword = Request.Cookies(ajlogin_cookiename)("strPassword")
			
			AttemptLogin strCookieUserName, strCookiePassword, true, false
		end if
	end if
	
	if CInt(Session("UserLevel")) < CInt(intTempPageLevel) then
		strQueryString = Request.QueryString
		if strQueryString <> "" then strQueryString = ("?" & Server.URLEncode(strQueryString))
		Response.Redirect (ajlogin_scripturl &"login.asp?m="& intTempPageLevel &"&u="& Request.ServerVariables("URL") & strQueryString)
	end if
End Function

Function AttemptLogin(ByVal strTempUserName, ByVal strTempPassword, ByVal boolTempUseCookies, ByVal boolTempHashPass)

	Set oMember = New Member
	if oMember.LoadBy("MemberNameUser", strTempUserName, "Text") then
		if (ajlogin_sendconfirmation AND oMember.MemberConfirmed = false) then
			AttemptLogin = "ERROR:UNCONFIRMED"
		elseif (ajlogin_requireapproval AND oMember.MemberApproved = false) then
			AttemptLogin = "ERROR:UNAPPROVED"
		else
			if boolTempHashPass then strTempPassword = Hash(strTempPassword)

			if strTempPassword = oMember.MemberPassword then
				Session.Timeout = 60	
				Session("UserIndex") = oMember.MemberIndex
				Session("UserName") = strTempUserName
				Session("UserLevel") = CInt(oMember.MemberAccessLevel)
				Session("UserEMail") = oMember.MemberEMail
				
				if boolTempUseCookies then
					Response.Cookies (ajlogin_cookiename)("strUserName") = strTempUserName
					Response.Cookies (ajlogin_cookiename)("strPassword") = strTempPassword
					Response.Cookies (ajlogin_cookiename).expires = DATE + 30 'Cookie expires in 30 days
				end if
				
				oMember.MemberDateLastLogin = Now()
				oMember.Save

				AttemptLogin = "SUCCESS"
			else
				AttemptLogin = "ERROR:BADLOGIN"
			end if
		end if
	else
		AttemptLogin = "ERROR:BADLOGIN" 'username not found
	end if
	
End Function

Function GenerateEMailBody(ByVal strTempMessageType, ByVal strTempEMail)
	Set oMember = New Member
	if oMember.LoadBy("MemberEMail", strTempEMail, "Text") then
	
		'Standard welcome message (used for body text of Welcome/Confirm messages)
		strEMailBody = "Welcome to "
		strEMailBody = strEMailBody & ajlogin_sitename & vbCrLf
		strEMailBody = strEMailBody & "--------------------------------" & vbCrLf
		strEMailBody = strEMailBody & "User Name: " & oMember.MemberNameUser & vbCrLf
		'strEMailBody = strEMailBody & "Password: " & oMember.MemberPassword & vbCrLf

		select case strTempMessageType
			case "CONFIRM"
				strEMailBody = vbCrLf & "Please visit the following URL to confirm this e-mail address for your "& ajlogin_sitename &" membership. "& VbCrLf & VbCrLf & ajlogin_scripturl &"confirm_email.asp?e="& Server.URLEncode(strTempEMail) &"&n="& oMember.MemberConfirmationToken & VbCrLf & VbCrLf & "Your confirmation token is " & oMember.MemberConfirmationToken

			case "PASSWORD"
				strEMailBody = "Automated password reset from " & ajlogin_sitename & " for " & oMember.MemberNameUser & "." & vbCrLf & vbCrLf
				strEMailBody = strEMailBody & "Please visit the following URL to reset your password."& VbCrLf & VbCrLf & ajlogin_scripturl &"reset_password.asp?e="& Server.URLEncode(strTempEMail) &"&n="& oMember.MemberPasswordResetToken & VbCrLf & VbCrLf & "Your password reset token is " & oMember.MemberPasswordResetToken
			case "NOTIFICATION"
				strEMailBody = "New Member Registration:" & vbCrLf
				strEMailBody = strEMailBody & "ID:" & oMember.MemberIndex & vbCrLf
				strEMailBody = strEMailBody & "User name:" & oMember.MemberNameUser & vbCrLf
				strEMailBody = strEMailBody & "E-mail address:" & oMember.MemberEMail & vbCrLf
				strEMailBody = strEMailBody & "Registration Date:" & oMember.MemberDateCreated & vbCrLf
		end select
	end if	
	
	GenerateEMailBody = strEMailBody
End Function

Function SendNotification(ByVal strTempEMail)
	SendNotification = SendEMail(ajlogin_emailnotification, ajlogin_emailfrom, ajlogin_emailreplyto, (ajlogin_sitename & " New Member Registration"), GenerateEMailBody("NOTIFICATION", strTempEMail), 1)
End Function

Function SendPassword(ByVal strTempEMail)
	SendPassword = SendEMail(strTempEMail, ajlogin_emailfrom, ajlogin_emailreplyto, (ajlogin_sitename & " Password Reset"), GenerateEMailBody("PASSWORD", strTempEMail), 1)
End Function

Function SendWelcome(ByVal strTempEMail)
	SendWelcome = SendEMail(strTempEMail, ajlogin_emailfrom, ajlogin_emailreplyto, ("Welcome to " & ajlogin_sitename), GenerateEMailBody("WELCOME", strTempEMail), 1)
End Function

Function SendConfirmation(ByVal strTempEMail)
	SendConfirmation = SendEMail(strTempEMail, ajlogin_emailfrom, ajlogin_emailreplyto, (ajlogin_sitename & " E-Mail Address Confirmation"), GenerateEMailBody("CONFIRM", strTempEMail), 1)
End Function

Function SendEMail(ByVal strTempTo, ByVal strTempFrom, ByVal strTempReplyTo, ByVal strTempSubject, ByVal strTempBody, ByVal intToMode)

	Dim objEmail
	
	if ajlogin_usecdonts then
		Set objEmail 		= CreateObject("CDONTS.NewMail")
		objEmail.Body    	= strTempBody
	else
		Set objEmail 		= CreateObject("CDO.Message")
		objEmail.ReplyTo 	= strTempReplyTo
		objEmail.TextBody  	= strTempBody
		With objEmail.Configuration.Fields
		
		            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		
		            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = ajlogin_cdosys_smtpserver
		
		            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = ajlogin_cdosys_sendusername
		
		            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ajlogin_cdosys_sendpassword
		
		            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = ajlogin_cdosys_smtpauthenticate
		
		            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = ajlogin_cdosys_smtpserverport
		
		            .Update
		
		End With
	end if
				
	select case CInt(intToMode)
		case 3
			objEmail.To		= strTempFrom
			objEmail.Cc 	= strTempTo
		case 2
			objEmail.To		= strTempFrom
			objEmail.Bcc 	= strTempTo
		case 1
			objEmail.To		= strTempTo
	end select

	objEmail.From = strTempFrom
	objEmail.Subject = strTempSubject

	On Error Resume Next
	
	objEmail.Send
	
	if Err <> 0 then
	   SendEMail = false
	else
	   SendEMail = true
	end if
			
	Set objEmail = Nothing

End Function

Function GetNA(ByVal strTempGetNA)
	if strTempGetNA = "" then
		GetNA = "n/a"
	else
		GetNA = strTempGetNA
	end if
End Function

Function GetSelected(ByVal strToReference, ByVal strToCompare)
	if strToReference <> strToCompare then
		GetSelected = ""
	else
		GetSelected = "selected"
	end if
End Function

Public Function Hash(ByVal strTempToHash)
	if ajlogin_shouldhashpasswords then
		Hash = sha256(strTempToHash & ajlogin_passwordsalt)
	else
		Hash = strTempToHash
	end if
End Function

Public Function GetMembers	

	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open ajlogin_dsn
	
	Set objRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM members"
		
	objRS.open strSQL, objConn

	Dim membersDictionary
	Set membersDictionary = CreateObject("Scripting.Dictionary")

	do until objRS.EOF
		Set tempMember = New Member
		tempMember.LoadBy "MemberIndex", objRS("MemberIndex"), "Number"
		membersDictionary.Add "" & objRS("MemberIndex"), tempMember
		Set tempMember = nothing
		objRS.MoveNext
	loop

	objRS.Close
	Set objRS = nothing		

	objConn.Close		
	Set objConn = nothing
	
	Set GetMembers = membersDictionary
	
End Function
%>

<script language="javascript" runat="server">
function RoundUp(x){
	return Math.ceil(x);
}

function GetUniqueIndex(){
	var millis = 1
	var date = new Date();
	var curDate = null;

	do { curDate = new Date(); }
	while(curDate-date < millis);

	return GetUnixTimeStamp().toString(16).toUpperCase();
}


function GetUnixTimeStamp(){
	var d = new Date();
	return d.getTime();
}
</script>

<%
Class MemberField
	Public FieldName
	Public FieldType
	Public FieldValue
	Public FormItem
End Class

Class MemberFormItem
	Public IsMemberEditable
	Public InputLabel
	Public InputName
	Public InputType
	Public InputClass
	Public InputValue
	Public Required
End Class

Class Member

	Public MemberIndex
	Public MemberAccessLevel
	Public MemberNameUser
	Public MemberPassword
	Public MemberEMail
	Public MemberNotes
	Public MemberConfirmed
	Public MemberConfirmationToken
	Public MemberPasswordResetToken
	Public MemberPasswordResetTokenDate
	Public MemberApproved
	Public MemberDateCreated
	Public MemberDateUpdated
	Public MemberDateLastLogin
	Public MemberFields
	
	Public Sub InitFields
	
		Set MemberFields = CreateObject("Scripting.Dictionary")

		Set oField = New MemberField
		oField.FieldName = "MemberNameFirst"
		oField.FieldType = "Text"
		Set oField.FormItem = New MemberFormItem
		oField.FormItem.IsMemberEditable = true
		oField.FormItem.InputLabel = "First Name"
		oField.FormItem.InputName = "MemberNameFirst"
		oField.FormItem.InputType = "text"
		oField.FormItem.InputClass = "text-input"
		oField.FormItem.Required = true 'whether or not to require the user to enter a value for this for item
		MemberFields.Add oField.FieldName, oField
		Set oField = nothing

		Set oField = New MemberField
		oField.FieldName = "MemberNameLast"
		oField.FieldType = "Text"
		Set oField.FormItem = New MemberFormItem
		oField.FormItem.IsMemberEditable = true
		oField.FormItem.InputLabel = "Last Name"
		oField.FormItem.InputName = "MemberNameLast"
		oField.FormItem.InputType = "text"
		oField.FormItem.InputClass = "text-input"
		oField.FormItem.Required = true 'whether or not to require the user to enter a value for this for item
		MemberFields.Add oField.FieldName, oField
		Set oField = nothing
		
	End Sub
	
	Public Function LoadBy(ByVal strProperty, ByVal strPropertyValue, ByVal strPropertyType)
	
		'Valid strPropertyType values are "Text", "Number" and "Date"
	
		InitFields
	
		Set objConn = Server.CreateObject("ADODB.Connection")
		Set objRS = Server.CreateObject("ADODB.Recordset")
	
		strValueWrapper = GetFieldWrapper(strPropertyType)
		
		strSQL = "SELECT * FROM members WHERE "& strProperty &" = " & strValueWrapper & Escape(strPropertyValue) & strValueWrapper
		
		objConn.Open ajlogin_dsn
		objRS.open strSQL, objConn
	
		if NOT objRS.EOF then
			MemberIndex = CInt(objRS("MemberIndex"))
			MemberAccessLevel = objRS("MemberAccessLevel")
			MemberNameUser = objRS("MemberNameUser")
			MemberPassword = objRS("MemberPassword")
			MemberEMail = objRS("MemberEMail")
			MemberNotes = objRS("MemberNotes")
			MemberConfirmed = CBool(objRS("MemberConfirmed"))
			MemberConfirmationToken = objRS("MemberConfirmationToken")
			MemberApproved = CBool(objRS("MemberApproved"))
			MemberDateCreated = objRS("MemberDateCreated")
			MemberDateUpdated = objRS("MemberDateUpdated")
			MemberDateLastLogin = objRS("MemberDateLastLogin")
			MemberPasswordResetToken = objRS("MemberPasswordResetToken")
			MemberPasswordResetTokenDate = objRS("MemberPasswordResetTokenDate")
			
			for each oField in MemberFields.Items
				oField.FieldValue = objRS(oField.FieldName)
			next
			
			LoadBy = true
		else
			LoadBy = false
		end if
		
		objRS.Close
		objConn.Close
		
		Set objConn = nothing
		Set objRS = nothing		
	End Function
	
	Public Sub SetFieldValue(ByVal strTempFieldName, ByVal strTempFieldValue)
		if MemberFields.Exists(strTempFieldName) then
			Set oField = MemberFields.Item(strTempFieldName)
			oField.FieldValue = strTempFieldValue
		end if
	End Sub
	
	Public Function Save
		Dim RecordsAffected
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open ajlogin_dsn
		
		intErrorId = 0
				
		Set objRS = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM members WHERE MemberEMail = '"& Escape(MemberEmail) &"'"
		if NOT IsEmpty(MemberIndex) then strSQL = strSQL & " AND MemberIndex <> "& MemberIndex
		objRS.open strSQL, objConn
		if NOT objRS.EOF then intErrorId = -1
		objRS.Close

		strSQL = "SELECT * FROM members WHERE MemberNameUser = '"& Escape(MemberNameUser) &"'"
		if NOT IsEmpty(MemberIndex) then strSQL = strSQL & " AND MemberIndex <> "& MemberIndex
		objRS.open strSQL, objConn
		if NOT objRS.EOF then intErrorId = -2
		objRS.Close
		Set objRS = nothing
		
		if intErrorId = 0 then
			if IsEmpty(MemberIndex) then
				strToken = GenToken(64)			
				strSQL = "INSERT INTO members(MemberPasswordResetToken, MemberPasswordResetTokenDate, MemberApproved, MemberConfirmed, MemberConfirmationToken, MemberAccessLevel, MemberNameUser, MemberPassword, MemberEMail, MemberNotes, MemberDateCreated, MemberDateUpdated, MemberDateLastLogin"
		
				for each oField in MemberFields.Items
					strSQL = strSQL & ", " & oField.FieldName
				next
				
				strSQL = strSQL & ") VALUES('', '"& ajlogin_pastdate &"', "&MemberApproved&", "&MemberConfirmed&", '"& strToken &"', "&MemberAccessLevel&", '"& Escape(MemberNameUser) &"', '"& Escape(MemberPassword) &"', '"& Escape(MemberEmail) &"', '"& Escape(MemberNotes) &"', '"& Now() &"', '"& Now() &"', '"& Now() &"'"
	
				for each oField in MemberFields.Items
					strFieldWrapper = GetFieldWrapper(oField.FieldType)
					strSQL = strSQL & ", " & strFieldWrapper & Escape(oField.FieldValue) & strFieldWrapper
				next
				
				strSQL = strSQL & ")"
			else
				strSQL = "UPDATE members SET MemberConfirmed = " & MemberConfirmed & ", "
				strSQL = strSQL & "MemberConfirmationToken = '" & MemberConfirmationToken & "', "
				strSQL = strSQL & "MemberAccessLevel = " & MemberAccessLevel & ", "
				strSQL = strSQL & "MemberNameUser = '" & Escape(MemberNameUser) & "', "
				strSQL = strSQL & "MemberPassword = '" & Escape(MemberPassword) & "', "
				strSQL = strSQL & "MemberEMail = '" & Escape(MemberEMail) & "', "
				strSQL = strSQL & "MemberNotes = '" & Escape(MemberNotes) & "', "
				strSQL = strSQL & "MemberDateCreated = '" & MemberDateCreated & "', "
				strSQL = strSQL & "MemberDateUpdated = '" & MemberDateUpdated & "', "
				strSQL = strSQL & "MemberDateLastLogin = '" & MemberDateLastLogin & "', "
				strSQL = strSQL & "MemberPasswordResetToken = '" & MemberPasswordResetToken & "', "
				strSQL = strSQL & "MemberPasswordResetTokenDate = '" & MemberPasswordResetTokenDate & "', "
		
				for each oField in MemberFields.Items
					strFieldWrapper = GetFieldWrapper(oField.FieldType)
					strSQL = strSQL & oField.FieldName & " = " & strFieldWrapper & Escape(oField.FieldValue) & strFieldWrapper & ", "
				next
		
				strSQL = strSQL & "MemberApproved = " & MemberApproved
				strSQL = strSQL & " WHERE MemberIndex = " & MemberIndex
			end if		
		
			objConn.execute strSQL, RecordsAffected
			Save = RecordsAffected
		else		
			Save = intErrorId
		end if
		
		objConn.Close		
		Set objConn = nothing
	End Function
	
	Public Function Escape(ByVal strToEscape)
		Escape = Replace(strToEscape, "'", "''")
	End Function
		
	Public Function GenToken(ByVal intLength)
		strTokenSource = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

		Randomize
		
		strToken = ""
		for i = 0 to intLength
			strTokenElementPosition = CInt(Len(strTokenSource)*Rnd+1) 'random int in range 1 to length of strTokenSource
			strToken = strToken & Mid(strTokenSource, strTokenElementPosition, 1)
		next
		
		GenToken = strToken	
	End Function
	
	Public Function GetFieldWrapper(ByVal strTempFieldType)
		strFieldWrapper = ""
		if(strTempFieldType = "Text") then strFieldWrapper = "'"
		if(strTempFieldType = "Date") then strFieldWrapper = "'"
		GetFieldWrapper = strFieldWrapper
	End Function

	Public Function Delete
		Dim RecordsAffected
		Set objConn = Server.CreateObject("ADODB.Connection")	
		objConn.Open ajlogin_dsn

		strSQL = "DELETE FROM members WHERE MemberIndex = " & MemberIndex		
		objConn.execute strSQL, RecordsAffected
		Delete = RecordsAffected
	
		objConn.Close		
		Set objConn = nothing	
	End Function	
	
	Public Function CheckPasswordResetToken(ByVal strToken)
		'24 hours = 1440 mins
		intMinsPassed = DateDiff("n", CDate(MemberPasswordResetTokenDate), Now)
		if MemberPasswordResetToken = strToken AND intMinsPassed < 1440 then 
			CheckPasswordResetToken = true
		else
			CheckPasswordResetToken = false
		end if
	End Function

	Public Function NewPasswordResetToken
		MemberPasswordResetToken = GenToken(64)
		MemberPasswordResetTokenDate = Now
		Save
	End Function

	Public Function ExpirePasswordResetToken
		MemberPasswordResetTokenDate = ajlogin_pastdate
		Save
	End Function
	
End Class

Class MemberForm
	Public EditMember
	Public Action
	Public Mode

	Public Sub RenderForm
		%>
		<script type="text/javascript">
			/*
			AJLogin uses the jQuery validation plugin to validate user submitted forms.
			
			For more information on how to add your own custom completion rules, see
			http://docs.jquery.com/Plugins/Validation/ 
			*/
			$(function() {
				<!--#INCLUDE file="inc_jsvars.asp"-->
				$("#member_form").validate(validation_options);
			});
		</script>
		<div class="member-form">
			<%
			if IsEmpty(EditMember) then
				Set EditMember = New Member
				EditMember.InitFields
			end if						
			%>
			<form id="member_form" method="post" action="<%=Action%>">
				<fieldset>
					<%
					if Mode = "ADMINNEW" OR Mode = "ADMINEDIT" OR Mode = "REGISTER" then
						%>
						<label for="MemberNameUser" class="label">User Name</label>
						<input type="text" name="MemberNameUser" id="MemberNameUser" class="text-input required" value="<%= Server.HTMLEncode(EditMember.MemberNameUser) %>" />
					<%
					end if
					
					if Mode = "ADMINNEW" OR Mode = "ADMINEDIT" then

						if EditMember.MemberConfirmed then
							strMemberConfirmedChecked = "checked=""checked"""
						else
							strMemberConfirmedChecked = ""
						end if

						if EditMember.MemberApproved then
							strMemberApprovedChecked = "checked=""checked"""
						else
							strMemberApprovedChecked = ""
						end if

						%>
						<label for="MemberAccessLevel" class="label">User Level</label>
						<input type="text" name="MemberAccessLevel" id="MemberAccessLevel" class="text-input required" value="<%= EditMember.MemberAccessLevel %>" />

						<div class="checkbox-group <% if NOT ajlogin_sendconfirmation then Response.Write("hide") %>">
							<input type="checkbox" name="MemberConfirmed" id="MemberConfirmed" class="text-input" value="1" <%= strMemberConfirmedChecked %> />
							<label for="MemberConfirmed" class="checkbox-label">E-mail Confirmed</label>
						</div>
						
						<div class="checkbox-group <% if NOT ajlogin_requireapproval then Response.Write("hide") %>">
							<input type="checkbox" name="MemberApproved" id="MemberApproved" class="text-input" value="1" <%= strMemberApprovedChecked %> />
							<label for="MemberApproved" class="checkbox-label">Account approved</label>
						</div>
						<%
					end if
	
					if Mode = "REGISTER" OR Mode = "ADMINNEW" then
						%>
						<label for="MemberPassword" class="label">Password</label>
						<input type="password" name="MemberPassword" id="MemberPassword" class="text-input required" value="" />
		
						<label for="MemberPassword2" class="label">Please re-enter password</label>
						<input type="password" name="MemberPassword2" id="MemberPassword2" class="text-input" value="" />
						<% 
					elseif Mode = "EDITPROFILE" then
						%><h2>Member Profile</h2><%
						Response.Write("<p><span class=""label"">User Name:</span> " & EditMember.MemberNameUser & "</p>")
						if EditMember.MemberAccessLevel > 1 then Response.Write("<p><span class=""label"">Access Level:</span> " & EditMember.MemberAccessLevel & "</p>")
					end if
					%>	
		
					<label for="MemberEMail" class="label">E-mail address</label>
					<input type="text" name="MemberEMail" id="MemberEMail" class="text-input email required" value="<%= Server.HTMLEncode(EditMember.MemberEMail) %>" />
					
					<%
					for each oField in EditMember.MemberFields.Items
						if(NOT IsEmpty(oField.FormItem)) then
							Set oFormItem = oField.FormItem
							if(oFormItem.IsMemberEditable OR Session("UserLevel") = ajlogin_adminlevel) then
								%>
								<label for="<%= oFormItem.InputName %>" class="label"><%= oFormItem.InputLabel %></label>
								<input type="<%= oFormItem.InputType %>" name="<%= oFormItem.InputName %>" id="<%= oFormItem.InputName %>" class="<%= oFormItem.InputClass %><% if oFormItem.Required then Response.Write(" required") %>" value="<%= Server.HTMLEncode(oField.FieldValue) %>" />
								<%
							end if
						end if
					next
					%>
					
					<%
					Select Case Mode
						Case "ADMINNEW"
							strButtonAction = "Create"
						Case "ADMINEDIT"
							strButtonAction = "Save"
						Case "REGISTER"
							strButtonAction = "Register"
						Case "EDITPROFILE"
							strButtonAction = "Save"
					End Select
					%>
	
					<input type="submit" name="member_form_submit" id="member_form_submit" class="submit-button" value="<%= strButtonAction %>" />
				</fieldset>
			</form>
			
			<% if Mode = "ADMINEDIT" OR Mode = "EDITPROFILE" then RenderChangePasswordForm %>

		</div>
		<%
			
	End Sub
	
	Public Sub RenderChangePasswordForm
		%>
		<script type="text/javascript">
			/*
			AJLogin uses the jQuery validation plugin to validate user submitted forms.
			
			For more information on how to add your own custom completion rules, see
			http://docs.jquery.com/Plugins/Validation/ 
			*/
			$(function() {
				<!--#INCLUDE file="inc_jsvars.asp"-->
				$("#change_pass_form").validate(validation_options);
			});
		</script>

		<div class="reset-password-form">
			<h2>Change Password</h2>
			<form id="change_pass_form" method="post" action="<%=Action%>">
				<fieldset>
					<% if Mode = "EDITPROFILE" then %>
						<label for="CurrentMemberPassword" class="label">Current Password</label>
						<input type="password" name="CurrentMemberPassword" id="CurrentMemberPassword" class="text-input required" value="" />
					<% elseif Mode = "RESETPASSWORD" then %>
						<input type="hidden" name="MemberEMail" id="MemberEMail" value="<%= Request.Querystring("e") %>" />
						<input type="hidden" name="MemberPasswordResetToken" id="MemberPasswordResetToken" value="<%= Request.Querystring("n") %>" />
					<% end if %>
		
					<label for="MemberPassword" class="label">New Password</label>
					<input type="password" name="MemberPassword" id="MemberPassword" class="text-input required" value="" />
		
					<label for="MemberPassword2" class="label">Please re-enter new password</label>
					<input type="password" name="MemberPassword2" id="MemberPassword2" class="text-input" value="" />
		
					<input type="submit" name="change_password_submit" id="change_password_submit" class="submit-button" value="Change Password" />
				</fieldset>
			</form>
		</div>
		<%
	End Sub
	
	Public Function SaveFormData
	
		if Mode = "REGISTER" and not ajlogin_enableregistration then
			SaveFormData = -3
			Exit Function
		end if

		if IsEmpty(EditMember) then
			Set EditMember = New Member
			EditMember.InitFields
		end if
		
		if Mode = "ADMINNEW" OR Mode = "ADMINEDIT" OR Mode = "REGISTER" then
			EditMember.MemberNameUser = Request.Form("MemberNameUser")
		end if

		if Mode = "ADMINNEW" OR Mode = "ADMINEDIT" then
			EditMember.MemberAccessLevel = Request.Form("MemberAccessLevel")
			
			if Request.Form("MemberConfirmed") <> "" then
				strMemberConfirmed = "TRUE"
			else 
				strMemberConfirmed = "FALSE"
			end if
			EditMember.MemberConfirmed = strMemberConfirmed
			
			if Request.Form("MemberApproved") <> "" then
				strMemberApproved = "TRUE"
			else 
				strMemberApproved = "FALSE"
			end if
			EditMember.MemberApproved = strMemberApproved
		end if
		
		if Mode = "ADMINNEW" OR Mode = "REGISTER" then
			EditMember.MemberPassword = Hash(Request.Form("MemberPassword"))
		end if

		if Mode = "REGISTER" then
			EditMember.MemberAccessLevel = "1"
			EditMember.MemberConfirmed = "FALSE"	
			EditMember.MemberApproved = "FALSE"	
		end if

		EditMember.MemberEMail = Request.Form("MemberEMail")

		for each oField in EditMember.MemberFields.Items
			if(NOT IsEmpty(oField.FormItem)) then
				Set oFormItem = oField.FormItem
				if(oFormItem.IsMemberEditable OR Session("UserLevel") = ajlogin_adminlevel) then
					FormData = Request.Form(oFormItem.InputName)
					EditMember.SetFieldValue oField.FieldName, FormData
				end if
			end if
		next
		
		SaveFormData = EditMember.Save
	End Function
	
	Public Function ChangePassword
		if NOT IsEmpty(EditMember) then
			strTempNewPass = Hash(Request.Form("MemberPassword"))
			
			if Mode = "EDITPROFILE" then
				strTempOldPass = Hash(Request.Form("CurrentMemberPassword"))
				if EditMember.MemberPassword = strTempOldPass then
					boolShouldContinue = true
				else
					boolShouldContinue = false
				end if
			elseif Mode = "RESETPASSWORD" then
				strToken = Request.Form("MemberPasswordResetToken")
				if EditMember.CheckPasswordResetToken(strToken) then
					boolShouldContinue = true
				else
					boolShouldContinue = false
				end if
			else
				boolShouldContinue = true
			end if
			
			if boolShouldContinue then
				EditMember.MemberPassword = strTempNewPass
				EditMember.Save
				if Mode = "RESETPASSWORD" then EditMember.ExpirePasswordResetToken
				ChangePassword = true
			else
				ChangePassword = false
			end if
		else
			ChangePassword = false
		end if
	End Function
	
	
End Class
%>

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
