<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>

<!--#INCLUDE file="inc_common.asp"-->
<% RestrictPageAccesToLevel(ajlogin_adminlevel) %>

<%
if Request.Form("bulk_action_indicies") <> "" then

	'Bulk Action Modes:
	' "DELETE"
	' "CONFIRM"
	' "UNCONFIRM"
	' "APPROVE"
	' "UNAPPROVE"

	strMode = Request.Form("bulk_action_mode")
	Set oMember = New Member
	arrIndecies = Split(Request.Form("bulk_action_indicies"), ",")
	for each i in arrIndecies

		if strMode = "DELETE" then
			oMember.MemberIndex = i
			oMember.Delete
		else
			oMember.LoadBy "MemberIndex", i, "Number"

			if strMode = "CONFIRM" then oMember.MemberConfirmed = True
			if strMode = "UNCONFIRM" then oMember.MemberConfirmed = False
	
			if strMode = "APPROVE" then oMember.MemberApproved = True
			if strMode = "UNAPPROVE" then oMember.MemberApproved = False

			oMember.Save
		end if		
	next
	Set oMember = nothing
end if
%>

<% PageTitle = "Admin Area" %>
<!--#INCLUDE file="inc_header.asp"-->

	<script type="text/javascript">
		/*
		AJLogin uses the jQuery tablesorter plugin to enhance the admin section.
		
		For information on how to customize this plugin, see http://tablesorter.com/docs/
		*/
		$(function() {
   			$("#member_table").tablesorter({
		        headers: { 
		            0: {
		                sorter: false
		            }
		        } 
    		});
			
			$(".icon-delete").click(function() {
				var deleteIndex = $(this).attr("id").substr(7) //id=delete_1
				if(confirm("Are you sure you want to delete member \""+ $("#"+deleteIndex).val() +"\"?")) {
					$("#bulk_action_mode").val("DELETE");
					$("#bulk_action_indicies").val(deleteIndex);
					$("#bulk_action_form").submit();
				}
			});

			$(".icon-approve").click(function() {
				var approveIndex = $(this).attr("id").substr(8) //id=approve_1
				$("#bulk_action_indicies").val(approveIndex);

				var isApproved = $(this).attr("alt");
				if(isApproved == "TRUE") {
					$("#bulk_action_mode").val("UNAPPROVE");
				} else {
					$("#bulk_action_mode").val("APPROVE");
				}

				$("#bulk_action_form").submit();
			});

			$(".icon-confirm").click(function() {
				var confirmIndex = $(this).attr("id").substr(8) //id=confirm_1
				$("#bulk_action_indicies").val(confirmIndex);

				var isConfirmed = $(this).attr("alt");
				if(isConfirmed == "TRUE") {
					$("#bulk_action_mode").val("UNCONFIRM");
				} else {
					$("#bulk_action_mode").val("CONFIRM");
				}
				
				$("#bulk_action_form").submit();
			});
			
			$(".admin-bulk-aciton").click(function(){
				var button_id = $(this).attr("id");
				var bulk_action_indicies = "", count = 0, str_prompt = "";
				$(".bulk-action-checkbox:visible:checked").each(function(){
					bulk_action_indicies += "," + $(this).attr("id");
					count++;
				});
				
				$("#bulk_action_form").attr('action', 'admin.asp');
				
				if(button_id == "admin_bulk_delete") {
					$("#bulk_action_mode").val("DELETE");
					str_prompt = "Are you sure you want to delete "+ count +" member"+ (count == 1 ? "" : "s") +"?";
				
				} else if(button_id == "admin_bulk_mail") {
					$("#bulk_action_form").attr('action', 'mass_email.asp');

				} else if(button_id == "admin_bulk_confirm") {
					$("#bulk_action_mode").val("CONFIRM");
					str_prompt = "Are you sure you want to mark "+ count +" member"+ (count == 1 ? "" : "s") +" e-mail addresses as confirmed?";
				
				} else if(button_id == "admin_bulk_unconfirm") {
					$("#bulk_action_mode").val("UNCONFIRM");
					str_prompt = "Are you sure you want to mark "+ count +" member"+ (count == 1 ? "" : "s") +" e-mail addresses as NOT confirmed?";
				
				} else if(button_id == "admin_bulk_approve") {
					$("#bulk_action_mode").val("APPROVE");
					str_prompt = "Are you sure you want to mark "+ count +" member"+ (count == 1 ? "" : "s") +" as approved?";
				
				} else if(button_id == "admin_bulk_unapprove") {
					$("#bulk_action_mode").val("UNAPPROVE");
					str_prompt = "Are you sure you want to mark "+ count +" member"+ (count == 1 ? "" : "s") +" as NOT approved?";
				}
				
				if(count < 1) {
					alert("No members selected. Check the box next to the edit button for each memeber you would like to select.");
				} else {					
					if((str_prompt != "" ? confirm(str_prompt) : true)) {
						$("#bulk_action_indicies").val(bulk_action_indicies.substr(1));
						$("#bulk_action_form").submit();
					}
				}
			});

			$("#filter").keyup(function () {
				filterUserList();
			});
			
			$("#select-all").click(function(){
				$(".bulk-action-checkbox:visible").each(function(){
					$(this).attr("checked", "checked");
					highlightRow($(this));
				});
			});

			$("#select-none").click(function(){
				$(".bulk-action-checkbox:visible").each(function(){
					$(this).attr("checked", "");
					highlightRow($(this));
				});
			});

			$("#admin-new-member-button").click(function (){
				document.location.href = '<%= ajlogin_scripturl %>edit_member.asp';
			});
			
			$(".filter-form-item").click(function(){
				filterUserList();
			});
			
			$(".bulk-action-checkbox").click(function(){
				highlightRow($(this));			
			});
			
			loadFilterFromCookie();
			filterUserList();
		});
		
		function highlightRow(o) {
			var bgcolor = "#FFFFFF";
			
			if(o.attr("checked") != "") {
				bgcolor = "#598EC9";
			}
			
			o.closest("tr").children().css("background-color", bgcolor);			
		}
		
		function loadFilterFromCookie() {	
			$("input[name='filter-confirmed']:checked").removeAttr("checked");
			
			//confirmed
			var confirmed = $.cookie('ajlogin_admin_filter_confirmed');
			if(confirmed == "" || confirmed == null) {
				$("#filter-confirmed-all").attr("checked", "checked");
			}
			else if(confirmed == ("1:<%=ajlogin_kConfirmed%>")) {
				$("#filter-confirmed-true").attr("checked", "checked");
			}
			else if(confirmed == ("0:<%=ajlogin_kConfirmed%>")) {
				$("#filter-confirmed-false").attr("checked", "checked");
			}

			//approved
			var approved = $.cookie('ajlogin_admin_filter_approved');
			if(approved == "" || approved == null) {
				$("#filter-approved-all").attr("checked", "checked");
			}
			else if(approved == ("1:<%=ajlogin_kApproved%>")) {
				$("#filter-approved-true").attr("checked", "checked");
			}
			else if(approved == ("0:<%=ajlogin_kApproved%>")) {
				$("#filter-approved-false").attr("checked", "checked");
			}

			//filter text
			var filterText = $.cookie('ajlogin_admin_filter_text');
			if(filterText != null) {
				$("#filter").val(filterText);
			}
		}
		
		function filterUserList() {
			var hidden;
		    var filter = $("#filter").val(), count = 0, total = 0;
			var confirmed = $("input[name='filter-confirmed']:checked").val();
			var approved = $("input[name='filter-approved']:checked").val();				
		    
		    $("#member_table tbody tr").each(function () {
		    	
		    	hidden = false;
		    	
		        if ($(this).text().search(new RegExp(filter, "i")) < 0) {
		            $(this).addClass("hide");
		            hidden = true;
		        }
		        		        
		        if ($(this).text().search(new RegExp(confirmed, "i")) < 0) {
		            $(this).addClass("hide");
		            hidden = true;
		        }
		        
		        if ($(this).text().search(new RegExp(approved, "i")) < 0) {
		            $(this).addClass("hide");
		            hidden = true;
				}

		        if(!hidden) {
		            $(this).removeClass("hide");
		            count++;
		        }
				
		        total++;
		    });
		    $("#filter-count").text(count);
		    $("#filter-count-total").text(total);

   		    //$("#persistent_filter").val([confirmed, approved );
		    
            $.cookie('ajlogin_admin_filter_confirmed', confirmed, { path: '/', expires: 10 });                    
            $.cookie('ajlogin_admin_filter_approved', approved, { path: '/', expires: 10 });                    
            $.cookie('ajlogin_admin_filter_text', filter, { path: '/', expires: 10 });                    
		}		
	</script>
	
	<div class="page-stripe">
		<p>You are logged in as <strong><%= Session("UserName") %></strong></p>
		<p>
			<a href="<%=ajlogin_scripturl%>members_area.asp">Member Area</a> | 
			<a href="<%=ajlogin_scripturl%>login.asp?m=LogOut">Log Out</a>
		</p>
		<div id="admin-new-member-div">
			<input type="button" id="admin-new-member-button" value="Create new member" />
		</div>
	</div>
			
	<div id="admin-left-column">		 
		<div id="filter-box">
			<h2>
				Filter members
				<span id="filter-count-wrapper">(showing <span id="filter-count">0</span> of <span id="filter-count-total">0</span>)</span>
			</h2>
			
			<div><input id="filter" type="text" name="filter" /></div>

			<div<% if NOT ajlogin_sendconfirmation then Response.Write(" class=""hide""") %>>
				<label for="filter-confirmed">E-mail confirmed</label>
				<span class="filter-form-item"><input id="filter-confirmed-all" type="radio" name="filter-confirmed" value="" checked="checked" /> <label for="filter-confirmed-all">all</label></span>
				<span class="filter-form-item"><input id="filter-confirmed-true" type="radio" name="filter-confirmed" value="1:<%=ajlogin_kConfirmed%>" /> <label for="filter-confirmed-true">yes</label></span>
				<span class="filter-form-item"><input id="filter-confirmed-false" type="radio" name="filter-confirmed" value="0:<%=ajlogin_kConfirmed%>" /> <label for="filter-confirmed-false">no</label></span>
			</div>
			
			<div<% if NOT ajlogin_requireapproval then Response.Write(" class=""hide""") %>>
				<label for="filter-approved">Account approved</label>
				<span class="filter-form-item"><input id="filter-approved-all" type="radio" name="filter-approved" value="" checked="checked" /> <label for="filter-approved-all">all</label></span>
				<span class="filter-form-item"><input id="filter-approved-true" type="radio" name="filter-approved" value="1:<%=ajlogin_kApproved%>" /> <label for="filter-approved-true">yes</label></span>
				<span class="filter-form-item"><input id="filter-approved-false" type="radio" name="filter-approved" value="0:<%=ajlogin_kApproved%>" /> <label for="filter-approved-false">no</label></span>
			</div>
		</div>
		
		<div id="admin_bulk_acitons_div">
			<h2>Bulk Actions</h2>
			<div id="bulk_delete_div" class="admin-bulk-aciton-buton"><img id="admin_bulk_delete" class="admin-bulk-aciton" src="images/delete.png" alt="Delete selected members" title="Delete selected members" /> Delete selected members<div class="ajlogin-clear"></div></div>
			<div id="bulk_mail_div" class="admin-bulk-aciton-buton"><img id="admin_bulk_mail" class="admin-bulk-aciton" src="images/mass-mail.png" alt="Mass E-mail" title="Send mass e-mail message to selected members" /> Send mass e-mail message to selected members<div class="ajlogin-clear"></div></div>
			<% if ajlogin_sendconfirmation then %>
				<div id="bulk_confirm_div" class="admin-bulk-aciton-buton"><img id="admin_bulk_confirm" class="admin-bulk-aciton" src="images/confirmed-true.png" alt="Mark e-mail address as confirmed" title="Mark selected e-mail address as confirmed" />Mark selected e-mail address as confirmed<div class="ajlogin-clear"></div></div>
				<div id="bulk_unconfirm_div" class="admin-bulk-aciton-buton"><img id="admin_bulk_unconfirm" class="admin-bulk-aciton" src="images/confirmed-false.png" alt="Mark e-mail address as NOT confirmed" title="Mark selected e-mail address as NOT confirmed" />Mark selected e-mail address as NOT confirmed<div class="ajlogin-clear"></div></div>
			<% end if %>
			<% if ajlogin_requireapproval then %>
				<div id="bulk_approve_div" class="admin-bulk-aciton-buton"><img id="admin_bulk_approve" class="admin-bulk-aciton" src="images/approved-true.png" alt="Mark account as approved" title="Mark selected accounts as approved" />Mark selected accounts as approved<div class="ajlogin-clear"></div></div>
				<div id="bulk_unapprove_div" class="admin-bulk-aciton-buton"><img id="admin_bulk_unapprove" class="admin-bulk-aciton" src="images/approved-false.png" alt="Mark account as NOT approved" title="Mark selected accounts as NOT approved" />Mark selected accounts as NOT approved<div class="ajlogin-clear"></div></div>
			<% end if %>			
		</div>
	</div>
	
	<div id="admin-right-column">
		<form id="bulk_action_form" method="post" action="<%=ajlogin_adminurl%>">
			<fieldset>
				<input type="hidden" id="bulk_action_mode" name="bulk_action_mode" value="" />
				<input type="hidden" id="bulk_action_indicies" name="bulk_action_indicies" value="" />
				<input type="hidden" id="persistent_filter" name="persistent_filter" value="" />
			</fieldset>
		</form>
		
		<table id="member_table" cellspacing="0">
			<thead>
				<tr>
					<th id="control"><span id="select-controls">Select<br /><span id="select-all">All</span> | <span id="select-none">None</span></span></th>
					<th id="MemberIndex"><span>ID</span></th>
					<th id="MemberNameUser"><span>User Name</span></th>

					<% if ajlogin_sendconfirmation then  %>
					<th id="MemberConfirmed">E-Mail<br />Confirmed</th>
					<% end if %>
					
					<% if ajlogin_requireapproval then  %>
					<th id="MemberApproved">Account<br />Approved</th>
					<% end if %>

					<th id="MemberAccessLevel"><span>Level</span></th>
					<th id="MemberEMail"><span>E-Mail</span></th>
					<th id="MemberNotes"><span>Notes</span></th>
					<th id="MemberDateCreated" class="date"><span>Created</span></th>
					<th id="MemberDateUpdated" class="date"><span>Modified</span></th>
					<th id="MemberDateLastLogin" class="date"><span>Last Log In</span></th>

					<%
					Set tempMember = New Member
					tempMember.InitFields
					
					for each columnName in tempMember.MemberFields.Keys
						Response.Write "<th id=""" & columnName & """><span>" & tempMember.MemberFields.Item(columnName).FormItem.InputLabel & "</span></th>"
					next

					Set tempMember = nothing
					%>
				</tr>
			</thead> 
			<tbody>
				<%
				Set membersDict = GetMembers
				
				for each oMember in membersDict.Items
					%>
					<tr>
						<td class="control">
							<input type="checkbox" id="<%= oMember.MemberIndex %>" class="bulk-action-checkbox" value="<%= oMember.MemberNameUser %>" />
							<a href="<%=ajlogin_scripturl%>edit_member.asp?id=<%= oMember.MemberIndex %>"><img src="images/edit.png" alt="Edit" class="icon icon-edit" /></a>
							<img id="delete_<%= oMember.MemberIndex %>" src="images/delete.png" alt="Delete" class="icon icon-delete" />
						</td>

						<td class="MemberIndex"><span><%= oMember.MemberIndex %></span></td>
						<td class="MemberNameUser"><span><%= oMember.MemberNameUser %></span></td>
	
						<%
						if oMember.MemberConfirmed then
							strMemberConfirmed = "TRUE"
							intMemberConfirmed = 1
						else
							strMemberConfirmed = "FALSE"
							intMemberConfirmed = 0
						end if

						if oMember.MemberApproved then
							strMemberApproved = "TRUE"
							intMemberApproved = 1
						else
							strMemberApproved = "FALSE"
							intMemberApproved = 0
						end if

						if ajlogin_sendconfirmation then Response.Write "<td class=""MemberConfirmed""><span class=""hide"">"& intMemberConfirmed & ":" & ajlogin_kConfirmed &"</span><img id=""confirm_"& oMember.MemberIndex &""" class=""icon icon-confirm"" src=""images/confirmed-"& LCase(strMemberConfirmed) &".png"" alt="""& strMemberConfirmed &""" /></td>"
						if ajlogin_requireapproval then Response.Write "<td class=""MemberApproved""><span class=""hide"">"& intMemberApproved & ":" & ajlogin_kApproved &"</span><img   id=""approve_"& oMember.MemberIndex &""" class=""icon icon-approve"" src=""images/approved-"& LCase(strMemberApproved) &".png"" alt="""& strMemberApproved &""" /></td>"
						%>
						
						<td class="MemberAccessLevel"><span><%= oMember.MemberAccessLevel %></span></td>
						<td class="MemberEMail"><span><%= oMember.MemberEMail %></span></td>
						<td class="MemberNotes"><span><%= Left(oMember.MemberNotes, 30) %></span></td>
						<td class="MemberDateCreated" class="date"><span><%= oMember.MemberDateCreated %></span></td>
						<td class="MemberDateUpdated" class="date"><span><%= oMember.MemberDateUpdated %></span></td>
						<td class="MemberDateLastLogin" class="date"><span><%= oMember.MemberDateLastLogin %></span></td>
	
						<%
						for each oField in oMember.MemberFields.Items
							Response.Write "<td class=""" & oField.FieldName & """><span>" & oField.FieldValue & "</span></td>"
						next
						%>
					</tr>
					<%	
				next
				%>
			</tbody> 
		</table>
	</div>		

<!--#INCLUDE file="inc_footer.asp"-->

<%
'AJLogin v4.1
'By: Adam Jarret
'RandomRavings.com
%>