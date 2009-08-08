var validation_options = {
	rules: {
		MemberPassword: {
			minlength: <%= ajlogin_minpasslength %>
		},
		MemberPassword2: {
			equalTo: "#MemberPassword"
		}
	},
	messages: {
		MemberPassword2: "Passwords did not match."
	}
};