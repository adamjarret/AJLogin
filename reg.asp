<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>

<!--#INCLUDE file="inc_common.asp"-->

<% PageTitle = "Please Register" %>
<!--#INCLUDE file="inc_header.asp"-->

	<% if strStatusMessage <> "" then %>
		<div class="status">
	    	<p><%=strStatusMessage%></p>
	    </div>
	<% end if %>
	
	<%
	Set oForm = New MemberForm
	oForm.Mode = "REGISTER"
	oForm.Action = ajlogin_scripturl & "reg.asp"
	
	if(Request.Form("member_form_submit") <> "") then
		strEmail = Request.Form("MemberEMail")
		intResult = CInt(oForm.SaveFormData)
		
		if intResult < 0  then
			select case intResult
				case -1
					strStatusMessage = "This E-mail address has already been registered. Do you need to <a href=""forgot_password.asp"">reset your password</a>?"
				case -2
					strStatusMessage = "This user name is unavailable."
			end select
			%>
			<div class="status">
		    	<p><%=strStatusMessage%></p>
		    </div>
			<%
		else		
			if ajlogin_sendconfirmation then
				SendConfirmation(strEmail)
			elseif ajlogin_sendwelcome then
				SendWelcome(strEmail)
			end if
			if ajlogin_sendnotification then SendNotification(strEmail)
			Response.Redirect(ajlogin_scripturl & "login.asp?m=New&d=" & Server.URLEncode(strEmail))
		end if
	end if

	oForm.RenderForm
	%>

	<div class="member-form">
		<br />
		Already a member? <a href="<%= ajlogin_scripturl %>login.asp">Click here to Log In.</a>
	</div>

<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>