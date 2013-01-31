<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
<!--#INCLUDE file="inc_common.asp"-->
<%
strStatusMessage = ""
strConfirmEmail = Request.QueryString("e")
strConfirmationToken = Request.QueryString("n")

if strConfirmEmail <> "" then
	Set oMember = New Member
	if oMember.LoadBy("MemberEMail", strConfirmEmail, "Text") then
		if oMember.MemberConfirmationToken = strConfirmationToken then
			oMember.MemberConfirmed = true
			oMember.Save
			Response.Redirect(ajlogin_scripturl & "login.asp?m=Confirmed&d=" & Server.URLEncode(strConfirmEmail))
		else
   			strStatusMessage = "Incorrect token for address " & strConfirmEmail
	   	end if
	else
   		strStatusMessage = "No user account found for " & strConfirmEmail
	end if
end if

%>

<% PageTitle = "Confirm E-mail Address" %>
<!--#INCLUDE file="inc_header.asp"-->

		<% if strStatusMessage <> "" then %>
			<div class="status">
		    	<p><%=strStatusMessage%></p>
		    </div>
		<% end if %>
		<div class="instructions-wrapper">
			<div class="instructions">
				<p><%=ajlogin_sitename%> sent a message to the e-mail address you provided upon registration to verify that you own the address.</p>
				<p>To confirm your address, please click the confirmation link in the message. Alternately, you can enter the token from the message and your e-mail address into the form below.</p>
				<p><a href="<%=ajlogin_scripturl%>resend_confirmation.asp">Click here to re-send the message</a>.</p>
			</div>
		</div>
		
		<div class="utility-form-div">
			<form class="utility-form" method="get" action="confirm_email.asp">
				<fieldset>
					<label for="e">E-mail address</label>
			        <input name="e" id="e" type="text" class="text-input email required" value="<%=strConfirmEmail%>" />
					
					<label for="n">Confirmation Token</label>
					<input name="n" id="n" type="text" class="text-input" value="<%=strConfirmationToken%>" />
			
					<input name="s" id="s" type="submit" value="Confirm Address" />
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