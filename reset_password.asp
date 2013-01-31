<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
<!--#INCLUDE file="inc_common.asp"-->

<% PageTitle = "Reset password" %>
<!--#INCLUDE file="inc_header.asp"-->

	<%	
	boolPromptForEmail  = true
	strStatusMessage = ""
	strEmail = Trim(Request.Querystring("e"))
	strToken = Trim(Request.Querystring("n"))
	
	if strEmail <> "" AND strToken <> "" then
		Set oForm = New MemberForm	
		oForm.Mode = "RESETPASSWORD"
		oForm.Action = ajlogin_scripturl &  "reset_password.asp"
		
		Set oForm.EditMember = New Member
		oForm.EditMember.LoadBy "MemberEMail", strEmail, "Text"
		
		if oForm.EditMember.CheckPasswordResetToken(strToken) then
			boolPromptForEmail  = false
			oForm.RenderChangePasswordForm
		else
			strStatusMessage = "Incorrect e-mail/token combination or token has expired."
		end if
	end if

	if(Request.Form("change_password_submit") <> "") then
		strEmail = Trim(Request.Form("MemberEMail"))

		Set oForm = New MemberForm	
		oForm.Mode = "RESETPASSWORD"
		oForm.Action = ajlogin_scripturl &  "reset_password.asp"
		
		Set oForm.EditMember = New Member
		oForm.EditMember.LoadBy "MemberEMail", strEmail, "Text"

		if oForm.ChangePassword then
			boolPromptForEmail  = false
			Response.Redirect(ajlogin_scripturl & "login.asp?m=ResetPass")
		else
			%><div class="status"><p>Incorrect e-mail/token combination or token has expired.</p></div><br /><%
		end if
	end if
	%>

	<% if strStatusMessage <> "" then %>
		<div class="status">
	    	<p><%=strStatusMessage%></p>
	    </div>
	<% end if %>
	
	<% if boolPromptForEmail then %>
		<div class="utility-form-div">
			<form class="utility-form" method="get" action="reset_password.asp">
				<fieldset>
					<label for="e">E-mail address</label>
					<input name="e" id="e" type="text" class="text-input email required" value="<%= strEmail %>" />
		
					<label for="n">Password reset token</label>
					<input name="n" id="n" type="text" class="text-input required" value="<%= strToken %>" />
			
					<input name="s" id="s" type="submit" value="Continue and set new password" />
				</fieldset>
			</form>
		</div>
	
		<div class="utility-form-bottom-links-div">	
			<a href="<%=ajlogin_scripturl%>login.asp">Return to Log In page</a>.
			<% if ajlogin_enableregistration then %>Need to <a href="<%=ajlogin_scripturl%>reg.asp">create an account</a>?<% end if %>
		</div>
	<% end if %>

<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>