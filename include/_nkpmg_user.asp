<%
'아이디가 없을 경우 메인 페이지 이동 처리(강제 접속 금지 처리) [허정호_20201124]
If IsEmpty(Request.Cookies("nkpmg_user")("coo_user_id")) Or IsNull(Request.Cookies("nkpmg_user")("coo_user_id")) Then
	Response.Write "<script type='text/javascript>'"
	Response.Write "alert('로그인 할수 없습니다. 관리자에게 문의 바랍니다....');"
	Response.Write "location.replace('./index.asp');"
	Response.Write "</script>"
Else
	user_id	= Request.Cookies("nkpmg_user")("coo_user_id")
	user_name = Request.Cookies("nkpmg_user")("coo_user_name")
	emp_no = Request.Cookies("nkpmg_user")("coo_emp_no")

	emp_company = Request.cookies("nkpmg_user")("coo_emp_company")
	saupbu = Request.cookies("nkpmg_user")("coo_saupbu")
	bonbu = Request.cookies("nkpmg_user")("coo_bonbu")
	team = Request.cookies("nkpmg_user")("coo_team")
	reside_place = Request.cookies("nkpmg_user")("coo_reside_place")
	reside_company = Request.cookies("nkpmg_user")("coo_reside_company")

	c_grade = Request.Cookies("nkpmg_user")("coo_grade")
	cost_grade = Request.Cookies("nkpmg_user")("coo_cost_grade")
	insa_grade = Request.Cookies("nkpmg_user")("coo_insa_grade")
	pay_grade = Request.Cookies("nkpmg_user")("coo_pay_grade")
	met_grade = Request.Cookies("nkpmg_user")("coo_met_grade")
	account_grade = Request.Cookies("nkpmg_user")("coo_account_grade")
	sales_grade = Request.Cookies("nkpmg_user")("coo_sales_grade")
	user_grade = Request.Cookies("nkpmg_user")("coo_user_grade")
	mg_group = Request.Cookies("nkpmg_user")("coo_mg_group")
	c_reside = Request.Cookies("nkpmg_user")("coo_reside")
	help_yn = Request.Cookies("nkpmg_user")("coo_help_yn")

	groupware_id = Request.Cookies("nkpmg_user")("coo_groupware_id")
	position = Request.Cookies("nkpmg_user")("coo_position")
	org_name = Request.Cookies("nkpmg_user")("coo_org_name")
	asset_company = Request.Cookies("nkpmg_user")("coo_asset_company")

	'중복 쿠키값 사용 확인
	reside = Request.Cookies("nkpmg_user")("coo_reside")
	c_name = Request.Cookies("nkpmg_user")("coo_user_name")

	'로그인 처리 시 사용 안된 쿠키 [허정호_20201124]
	'hp = Request.Cookies("nkpmg_user")("coo_hp") = rs("hp")
End If

'Response.write  "c_grade :"         & c_grade         & "<br>"
'Response.write  "cost_grade :"      & cost_grade      & "<br>"
'Response.write  "insa_grade :"      & insa_grade      & "<br>"
'Response.write  "pay_grade :"       & pay_grade       & "<br>"
'Response.write  "met_grade :"       & met_grade       & "<br>"
'Response.write  "account_grade :"   & account_grade   & "<br>"
'Response.write  "sales_grade :"     & sales_grade     & "<br>"
'Response.write  "c_reside :"        & c_reside        & "<br>"
'Response.write  "user_id :"         & user_id         & "<br>"
'Response.write  "user_name :"       & user_name       & "<br>"
'Response.write  "user_grade :"      & user_grade      & "<br>"
'Response.write  "reside_place :"    & reside_place    & "<br>"
'Response.write  "reside_company :"  & reside_company  & "<br>"
'Response.write  "reside :"          & reside          & "<br>"
'Response.write  "mg_group :"        & mg_group        & "<br>"
'Response.write  "help_yn :"         & help_yn         & "<br>"
'Response.write  "asset_company :"   & asset_company   & "<br>"
'Response.write  "c_name :"          & c_name          & "<br>"
'Response.write  "emp_no :"          & emp_no          & "<br>"
'Response.write  "bonbu :"           & bonbu           & "<br>"
'Response.write  "saupbu :"          & saupbu          & "<br>"
'Response.write  "team :"            & team            & "<br>"
'Response.write  "org_name :"        & org_name        & "<br>"
'Response.write  "position :"        & position        & "<br>"
'Response.write  "emp_company :"     & emp_company     & "<br>"
'Response.write  "groupware_id :"    & groupware_id    & "<br>"

%>
