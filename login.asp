<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
<!--#INCLUDE file="inc_common.asp"-->
<%
strStatusMessage = ""
strAttemptedURL = Request.QueryString("u")
strPageMode = Request.QueryString("m")
strPageData = Request.QueryString("d")

if IsNumeric(strPageMode) and strPageMode<>"" then
  intRequiredLevel = strPageMode 
  strPageMode = "InsufficientAccess"
end if

if strPageMode = "LogIn" then
	if Request.Form("UseCookies") <> "" then
		UseCookies = true
	else
		UseCookies = false
	end if
	
	strLoginResult = AttemptLogin(Request.Form("UserName"), Request.Form("Password"), UseCookies, true)
	
	if strLoginResult = "SUCCESS" then
		if Request.Form("AttemptedURL") <> "" then
			Response.Redirect Request.Form("AttemptedURL")
		else
			if CInt(Session("UserLevel")) >= ajlogin_adminlevel then
				Response.Redirect ajlogin_adminurl
			else
				Response.Redirect ajlogin_memberurl
			end if
		end if
	else
		if strLoginResult = "ERROR:UNCONFIRMED" then
			strPageMode = "ConfirmEMail"
		elseif strLoginResult = "ERROR:UNAPPROVED" then
			strPageMode = "ApproveMember"
		else
			strPageMode = "BadLogin"
		end if

	end if
end if

select case strPageMode
 	case "BadLogin"
   		strStatusMessage = "Incorrect User Name/Password"
 	case "InsufficientAccess"
		if CInt(intRequiredLevel) = 1 then
     		strStatusMessage="You must be logged-in to view this page."
   		else
     		strStatusMessage="Level "& intRequiredLevel &" access is required to view this page."
   		end if
 	case "ApproveMember"
     	strStatusMessage="You account has not been approved by a site admin. You will not be able to log in until your account is approved."
 	case "ConfirmEMail"
     	strStatusMessage="The e-mail address for this account has not been confirmed.<br /><a href="""& ajlogin_scripturl &"confirm_email.asp"">Click here to confirm your address</a>."
 	case "Confirmed"
     	strStatusMessage="Your e-mail address has been verified."
    case "SendPass"
     	strStatusMessage="A password reset link has been sent to " & strPageData &". The link will expire in 24 hours."
    case "ResetPass"
     	strStatusMessage="Your password has been changed."
    case "SendConf"
     	strStatusMessage="An address ownership verification message has been sent to " & strPageData  	
    case "New"
		if ajlogin_sendconfirmation then strStatusMessage = strStatusMessage & "<p>An address ownership verification message has been sent to " & strPageData & ". Please confirm your e-mail address by clicking the confirmation link in this message.</p>"
    	if ajlogin_requireapproval then strStatusMessage = strStatusMessage & "<p>You will not be able to Log In until your account is approved by a site admin.</p>"
 	case "LogOut"
   		Response.Cookies (ajlogin_cookiename) = ""
   		Session.Abandon
   		strStatusMessage="You Are Now Logged Out."
end select
%>

<% PageTitle = "Please Log In" %>
<!--#INCLUDE file="inc_header.asp"-->

		<% if strStatusMessage <> "" then %>
			<div class="status">
		    	<p><%=strStatusMessage%></p>
		    </div>
		<% end if %>
		<div class="utility-form-div">
			<form class="utility-form" method="post" action="login.asp?m=LogIn">
				<fieldset>
					<input type="hidden" name="AttemptedURL" id="AttemptedURL" value="<%=strAttemptedURL%>" />
					
					<label for="UserName">User Name</label>
	      			<input name="UserName" id="UserName" type="text" class="text-input required" value="<%=strEMailAddress%>" />
					
					<label for="Password">Password</label>
	      			<input name="Password" id="Password" type="password" class="text-input required" value="<%=intConfirmationCode%>" />
	      			
	      			<div class="checkbox-group">
		      			<input name="UseCookies" id="UseCookies" type="checkbox" value="1" title="Requires cookies to be enabled." />
						<label for="UseCookies" title="Requires cookies to be enabled.">Keep me logged-in on this computer</label>
					</div>
					
					<input name="s" type="submit" value="Log In" />
				</fieldset>
  			</form>
  		</div>
  		
		<div class="utility-form-bottom-links-div">
			<% if ajlogin_sendpassword then %>
			<a href="<%=ajlogin_scripturl%>forgot_password.asp">Forgot your password</a>?
			<% end if %>
			Need to <a href="<%=ajlogin_scripturl%>reg.asp">create an account</a>?
		</div>


<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>