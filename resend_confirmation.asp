<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
<!--#INCLUDE file="inc_common.asp"-->
<%
strStatusMessage = ""
strPageMode = Request.QueryString("m")

if strPageMode = "SendConf" then
	strEmail = Trim(Request.Form("EMail"))
	
	Set oMember = New Member
	if oMember.LoadBy("MemberEMail", strEmail, "Text") then
		if SendConfirmation(strEmail) then
			Response.Redirect(ajlogin_scripturl & "login.asp?m=SendConf&d=" & Server.URLEncode(strEmail))
		else
   			strStatusMessage = "Error sending message to " & strEmail
	   	end if
	else
   		strStatusMessage = "No user account found for " & strEmail
	end if
end if

%>

<% PageTitle = "Re-send E-mail Confirmation Message" %>
<!--#INCLUDE file="inc_header.asp"-->

		<% if strStatusMessage <> "" then %>
			<div class="status">
		    	<p><%=strStatusMessage%></p>
		    </div>
		<% end if %>
		<div class="utility-form-div">
			<form class="utility-form" method="post" action="resend_confirmation.asp?m=SendConf">
				<fieldset>
	    			<label for="EMail">E-mail address</label>
	 				<input name="EMail" id="EMail" type="text" class="text-input email required" value="<%=strEmail%>" />
	
					<input name="s" id="s" type="submit" value="Send Message" />
				</fieldset>
  			</form>
		</div>
		
		<div class="utility-form-bottom-links-div">
	    	Return to <a href="<%=ajlogin_scripturl%>login.asp">Log In</a> page.
	    	<% if ajlogin_enableregistration then %>Need to <a href="<%=ajlogin_scripturl%>reg.asp">create an account</a>?<% end if %>
		</div>

<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>