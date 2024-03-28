<%
'계정 인증 체크
If Request.Cookies("nkpmg_user")("coo_user_id") = "" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	location.replace('warning.asp');"
	Response.Write "</script>"
    Response.End
End If

Dim coo_grade, coo_cost_grade, coo_insa_grade, coo_pay_grade, coo_met_grade
Dim coo_account_grade, coo_sales_grade, coo_reside, coo_user_id, coo_user_name
Dim coo_user_grade, coo_reside_place, coo_reside_company, coo_mg_group, coo_help_yn
Dim coo_asset_company
Dim coo_emp_no	'직원 사번
Dim coo_bonbu, coo_saupbu, coo_team
Dim coo_org_name, coo_position, coo_emp_company, coo_groupware_id

'설정된 쿠키 값 변수 저장
coo_grade        = Request.Cookies("nkpmg_user")("coo_grade")
coo_cost_grade     = Request.Cookies("nkpmg_user")("coo_cost_grade")
coo_insa_grade     = Request.Cookies("nkpmg_user")("coo_insa_grade")
coo_pay_grade      = Request.Cookies("nkpmg_user")("coo_pay_grade")
coo_met_grade      = Request.Cookies("nkpmg_user")("coo_met_grade")
coo_account_grade  = Request.Cookies("nkpmg_user")("coo_account_grade")
coo_sales_grade    = Request.Cookies("nkpmg_user")("coo_sales_grade")
coo_reside       = Request.Cookies("nkpmg_user")("coo_reside")
coo_user_id        = Request.Cookies("nkpmg_user")("coo_user_id")
coo_user_name      = Request.Cookies("nkpmg_user")("coo_user_name")
coo_user_grade     = Request.Cookies("nkpmg_user")("coo_user_grade")
coo_reside_place   = Request.Cookies("nkpmg_user")("coo_reside_place")
coo_reside_company = Request.Cookies("nkpmg_user")("coo_reside_company")
coo_mg_group       = Request.Cookies("nkpmg_user")("coo_mg_group")
coo_help_yn        = Request.Cookies("nkpmg_user")("coo_help_yn")
coo_asset_company  = Request.Cookies("nkpmg_user")("coo_asset_company")
coo_emp_no         = Request.Cookies("nkpmg_user")("coo_emp_no")
coo_bonbu          = Request.Cookies("nkpmg_user")("coo_bonbu")
coo_saupbu         = Request.Cookies("nkpmg_user")("coo_saupbu")
coo_team           = Request.Cookies("nkpmg_user")("coo_team")
coo_org_name       = Request.Cookies("nkpmg_user")("coo_org_name")
coo_position       = Request.Cookies("nkpmg_user")("coo_position")
coo_emp_company    = Request.Cookies("nkpmg_user")("coo_emp_company")
coo_groupware_id   = Request.Cookies("nkpmg_user")("coo_groupware_id")
%>