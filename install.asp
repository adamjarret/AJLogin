<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>

<!--#INCLUDE file="inc_common.asp"-->
<% PageTitle = "Install AJLogin 4.1" %>
<!--#INCLUDE file="inc_header.asp"-->

	<div class="page-stripe">
		<p>Create your administrator account.</p>
	</div>
	
	<div class="ajlogin-install">
		
		<%
		Set oForm = New MemberForm
		oForm.Action = ajlogin_scripturl & "install.asp"
		oForm.Mode = "ADMINNEW"
		Set oForm.EditMember = New Member
		oForm.EditMember.InitFields
		oForm.EditMember.MemberAccessLevel = 3
		oForm.EditMember.MemberConfirmed = true
		oForm.EditMember.MemberApproved = true
		
		if(Request.Form("member_form_submit") <> "") then
	
			intResult = CInt(oForm.SaveFormData)
			
			if intResult < 0  then
				%>
				<div class="status">
			    	<p>Error creating admin account.</p>
			    </div>
				<%
			else
				AttemptLogin Request.Form("MemberNameUser"), Request.Form("MemberPassword"), false, true
				Response.Redirect ajlogin_adminurl 
			end if
	
		end if
	
		oForm.RenderForm
		%>
	
	</div>
<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>