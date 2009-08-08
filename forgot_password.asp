<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
<!--#INCLUDE file="inc_common.asp"-->
<%
strStatusMessage = ""
strPageMode = Request.QueryString("m")

if strPageMode = "SendPass" then
	strEmail = Trim(Request.Form("EMail"))
	
	Dim oMember
	Set oMember = New Member
	if oMember.LoadBy("MemberEMail", strEmail, "Text") then
		oMember.NewPasswordResetToken
		if SendPassword(strEmail) then
			Response.Redirect(ajlogin_scripturl & "login.asp?m=SendPass&d=" & Server.URLEncode(strEmail))
		else
   			strStatusMessage = "Error sending message to " & strEmail
	   	end if
	else
   		strStatusMessage = "No user account found for " & strEmail
	end if
end if
%>

<% PageTitle = "Send Password Reset Link" %>
<!--#INCLUDE file="inc_header.asp"-->
		<% if strStatusMessage <> "" then %>
			<div class="status">
		    	<p><%=strStatusMessage%></p>
		    </div>
		<% end if %>
		<div class="utility-form-div">
			<form class="utility-form" method="post" action="forgot_password.asp?m=SendPass">
				<fieldset>
	  				<label for="EMail">E-mail address</label>
	      			<input name="EMail" id="EMail" type="text" class="text-input email required" value="<%=strEmail%>" />
	
					<input name="s" id="s" type="submit" value="Send Link" />
				</fieldset>
  			</form>
		</div>
		
		<div class="utility-form-bottom-links-div">
			<a href="<%=ajlogin_scripturl%>login.asp">Return to Log In page</a>.
	    	Need to <a href="<%=ajlogin_scripturl%>reg.asp">create an account</a>?
		</div>
<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>