<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>

<!--#INCLUDE file="inc_common.asp"-->
<% RestrictPageAccesToLevel(1) %>

<%
 PageTitle = "Edit Profile"
 strPageMode = Request.QueryString("m")
%>
<!--#INCLUDE file="inc_header.asp"-->

	<% if strStatusMessage <> "" then %>
		<div class="status">
			<p><%=strStatusMessage%></p>
		</div>
	<% end if %>
	
	<%
	Set oForm = New MemberForm	
	oForm.Mode = "EDITPROFILE"
	oForm.Action = ajlogin_scripturl &  "edit_profile.asp"
	
	Set oForm.EditMember = New Member
	oForm.EditMember.LoadBy "MemberIndex", Session("UserIndex"), "Number"
	
	if(Request.Form("member_form_submit") <> "") then
		intResult = CInt(oForm.SaveFormData)
		
		if intResult < 0  then
			select case intResult
				case -1
					strStatusMessage = "This e-mail address has already been registered to another account. Do you need to <a href=""forgot_password.asp"">reset the password</a> to that account?"
				case -2
					strStatusMessage = "This user name is unavailable."
				case -3
					strStatusMessage = "User registration has been disabled."
			end select
		else		
			strStatusMessage = "Your profile has been updated."
		end if
		%>
		<div class="status">
	    	<p><%=strStatusMessage%></p>
	    </div>
		<%
	end if

	if(Request.Form("change_password_submit") <> "") then
		if oForm.ChangePassword then
			%><div class="status"><p>Your password has been changed.</p></div><br /><%
		else
			%><div class="status"><p>Error: Incorrect current password.</p></div><br /><%
		end if
	end if

	oForm.RenderForm
	%>

	<div class="member-form">
		<br />
		<a href="<%= ajlogin_memberurl %>">Back to Member Area</a> | 
		<a href="<%= ajlogin_scripturl %>login.asp?m=LogOut">Log Out</a>
	</div>

<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>