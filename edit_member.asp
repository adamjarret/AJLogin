<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>

<!--#INCLUDE file="inc_common.asp"-->
<% RestrictPageAccesToLevel(3) %>

<%
 strEditMemberIndex = Request.QueryString("id")
 if strEditMemberIndex <> "" then
 	PageTitle = "Edit Member"
 else
	 PageTitle = "New Member"
 end if
%>
<!--#INCLUDE file="inc_header.asp"-->

	<div class="page-stripe">
		<p>You are logged in as <strong><%= Session("UserName") %></strong></p>
		<p>
			<a href="<%=ajlogin_memberurl%>">Member area</a> | 
			<a href="<%=ajlogin_scripturl%>login.asp?m=LogOut">Log Out</a>
		</p>
		<p><a href="<%=ajlogin_adminurl%>">Back to Admin area</a></p>
	</div>	

	<%
	Set oForm = New MemberForm
	oForm.Action = ajlogin_scripturl & "edit_member.asp?id=" & strEditMemberIndex
	
	if strEditMemberIndex <> "" then
		oForm.Mode = "ADMINEDIT"
		Set oForm.EditMember = New Member
		oForm.EditMember.LoadBy "MemberIndex", strEditMemberIndex, "Number"
	else
		oForm.Mode = "ADMINNEW"
	end if
	
	if(Request.Form("member_form_submit") <> "") then

		intResult = CInt(oForm.SaveFormData)
		
		if intResult < 0  then
			select case intResult
				case -1
					strStatusMessage = "This E-mail address has already been registered."
				case -2
					strStatusMessage = "This user name is unavailable."
			end select
		else		
			if oForm.Mode = "ADMINNEW" then
				Response.Redirect(ajlogin_adminurl)
			else
				strStatusMessage = "Member profile has been updated."
			end if
		end if

		%>
		<div class="status">
	    	<p><%=strStatusMessage%></p>
	    </div>
		<%
	end if

	if(Request.Form("change_password_submit") <> "") then
		if oForm.ChangePassword then
			%><div class="status"><p>Member password has been changed.</p></div><br /><%
		else
			%><div class="status"><p>Error: New password cannot be saved.</p></div><br /><%
		end if
	end if

	oForm.RenderForm
	%>
	
<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>