<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>
<!--#INCLUDE file="inc_common.asp"-->
<% RestrictPageAccesToLevel(ajlogin_adminlevel) %>

<% PageTitle = "Send Mass E-mail" %>
<!--#INCLUDE file="inc_header.asp"-->

	<script type="text/javascript">
		$(function() {
			$("#why_to").click(function(){
				alert('When CC or BCC is selected the default FROM address is used as the TO address.\n\nWhen the recipients are added to the BCC field, each reciepient does not see the full list of e-mail addresses attached to the message.');
			});
			
			$("#bcc_toggle").change(function(){
				if($(this).val() == "1") {
					$("#bcc_label").text('To');
					$("#to_field").addClass('hide');
				} else {
					$("#bcc_label").text(($(this).val() == "2" ? "BCC" : "CC" ));
					$("#to_field").removeClass('hide');
				}
			});
			
			$("#bcc_toggle").change();
		});
	</script>

	<%	
	strPageMode = "SEND"
	strStatusMessage = ""
	strAddresses = ""
	strIndicies = Request.Form("bulk_action_indicies")
	
	if Request.Form("send_mail_action") <> "" then
		if SendEMail(Request.Form("bcc"), ajlogin_emailfrom, ajlogin_emailreplyto, Request.Form("sub"), Request.Form("b"), Request.Form("bcc_toggle")) then
			strStatusMessage = "Message Sent"
			strPageMode = "SENT"
		else
			strStatusMessage = "Error Sending Message"
		end if
		strAddresses = Request.Form("bcc")
	else
		if strIndicies <> "" then
			Set oMember = New Member
			arrIndecies = Split(strIndicies, ",")
			for each i in arrIndecies
					oMember.LoadBy "MemberIndex", i, "Number"
					strAddresses = strAddresses & ";" & oMember.MemberEMail
			next
			strAddresses = Right(strAddresses, Len(strAddresses)-1)
		else
			strStatusMessage = "No recipients specified."
			strPageMode = "ERROR"
		end if
	end if
	%>
	
	<% if strStatusMessage <> "" then %>
		<div class="status">
	    	<p><%=strStatusMessage%></p>
	    </div>
	<% end if %>	

	<% if strPageMode <> "ERROR" then %>
		<div class="utility-form-div">
			<form class="utility-form" method="post" action="mass_email.asp">
				<fieldset>
				
					<br />
				
					Recipient e-mail addresses will be added to the <select name="bcc_toggle" id="bcc_toggle" <% if strPageMode <> "SEND" then Response.Write("disabled=""disabled""") %>>
						<option value="2"<% if Request.Form("bcc_toggle") = "2" OR Request.Form("bcc_toggle") = "" then Response.Write("selected=""selected""")%>>BCC</option>
						<option value="3"<% if Request.Form("bcc_toggle") = "3" then Response.Write("selected=""selected""")%>>CC</option>
						<option value="1"<% if Request.Form("bcc_toggle") = "1" then Response.Write("selected=""selected""")%>>To</option>
					</select> field.
				
					<div id="to_field">
						<label for="">To</label>
						<%=ajlogin_emailfrom%>
						<a id="why_to" href="#">Why is the message addressed to me?</a>
					</div>
						
					<label id="bcc_label" for="bcc">BCC</label>
					<%= Replace(strAddresses, ";", ", ") %>
					<input name="bcc" id="bcc" type="hidden" value="<%= strAddresses %>" />
		
					<label for="sub">Subject</label>
					<% if strPageMode <> "SENT" then %>
						<input name="sub" id="sub" type="text" class="text-input" value="<%= Request.Form("sub") %>" />
					<% else %>
						<%= Request.Form("sub") %>
					<% end if %>
	
					<label for="b">Body</label>
					<% if strPageMode <> "SENT" then %>
						<textarea name="b" id="b" class="text-input required"><%= Request.Form("b") %></textarea>
					<% else %>
						<%= Replace(Request.Form("b"), VbCrLf, "<br />") %>
					<% end if %>
			
					<% if strPageMode <> "SENT" then %>
						<input name="send_mail_action" id="send_mail_action" type="submit" value="Send Message" />
					<% end if %>
				</fieldset>
			</form>
		</div>
	
	<% end if %>

	<div class="utility-form-bottom-links-div">	
		<p><a href="<%=ajlogin_scripturl%>admin.asp"><% if strPageMode <> "SENT" then Response.Write("Cancel and ") %>Return to the Admin page</a></p>
	</div>

<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>