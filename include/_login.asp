<!--#include virtual = "/common/inc_top.asp"--><!--설정 파일-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<%
'==========================
'author : 허정호
'modify date : 20201117
'Desc :
'	설정 파일 include 추가
'	변수 선언 추가 및 사용 객체 소멸 처리
'==========================

	Dim id
	Dim pass
	Dim save_id
	Dim DBConn, rs, sql
	Dim login_cnt

	id = Request.Form("id")
	pass = Request.Form("pass")
	save_id = Request.Form("save_id")

	If id = "" Or IsNull(id) Or id = " " Or (id > "700000" And id < "799999") Then
		Response.write"<script language=javascript>"
		Response.write"alert('아이디가 등록되어 있지 않습니다....');"
		Response.write"history.go(-1);"
		Response.write"</script>"
	End If

	Set DBConn = Server.CreateObject("ADODB.CONNECTION")
	DBConn.Open DbConnect

    ' 개행문자 제거 박영주 요청 2019-06-04
    ' 가끔 들어가는 경우가 있어 아예 로그인시 전체처리.. as등록시 작업자 추가에서 오류가 남..
    sql="UPDATE memb SET org_name = REPLACE(org_name, '\r', '') WHERE org_name LIKE '%\r%' "
    DBConn.Execute(sql)

    sql="UPDATE memb SET org_name = REPLACE(org_name, '\n', '') WHERE org_name LIKE '%\n%' "
    DBConn.Execute(sql)
    ' 개행문자 제거 박영주 요청 2019-06-04

	'조회 컬럼 표기[허정호_20201117]
	'sql="select * from memb where user_id='"&id&"'"
	sql = "SELECT user_id, user_name, emp_no, pass, grade, emp_company, saupbu, bonbu, " &_
		"team, reside, reside_place, reside_company, hp, cost_grade, " &_
		"insa_grade, pay_grade, met_grade, account_grade, sales_grade, " &_
		"user_grade, mg_group, help_yn, email, position, " &_
		"org_name, asset_company, login_cnt " &_
		"FROM memb " &_
		"WHERE user_id = '" & id &"' "

	Set rs = DBConn.Execute(sql)

	If rs.EOF or rs.BOF Then
		Response.write "<script language=javascript>"
		Response.write "	alert('아이디가 등록되어 있지 않습니다....');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	ElseIf rs("pass") <> pass Then
		Response.write "<script language=javascript>"
		Response.write "	alert('비밀번호가 다릅니다....');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	ElseIf rs("grade") = "6" Then
		Response.write "<script language=javascript>"
		Response.write "	alert('NKP 사용권한이 없거나 조직변경으로 사번이 바뀔수 있습니다.');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	ElseIf rs("grade") = "" Or IsNull(rs("grade")) Then
		Response.write "<script language=javascript>"
		Response.write "	alert('로그인 할수 없습니다. 관리자에게 문의 바랍니다....');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	Else
		Response.Cookies("nkpmg_user")("coo_user_id") = rs("user_id")
		Response.Cookies("nkpmg_user")("coo_user_name") = rs("user_name")
		Response.Cookies("nkpmg_user")("coo_emp_no") = rs("emp_no")

		If IsNull(rs("emp_company")) Then
			Response.Cookies("nkpmg_user")("coo_emp_company") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_emp_company") = rs("emp_company")
		End If

		If IsNull(rs("saupbu")) Then
			Response.Cookies("nkpmg_user")("coo_saupbu") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_saupbu") = rs("saupbu")
		End If

		If IsNull(rs("bonbu")) Then
			Response.Cookies("nkpmg_user")("coo_bonbu") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_bonbu") = rs("bonbu")
		End If

		If IsNull(rs("team")) Then
			Response.Cookies("nkpmg_user")("coo_team") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_team") = rs("team")
		End If

		If IsNull(rs("reside_place")) Then
			Response.Cookies("nkpmg_user")("coo_reside_place") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_reside_place") = rs("reside_place")
		End If

		If IsNull(rs("reside_company")) Then
			Response.Cookies("nkpmg_user")("coo_reside_company") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_reside_company") = rs("reside_company")
		End If

		Response.Cookies("nkpmg_user")("coo_hp") = rs("hp")
		Response.Cookies("nkpmg_user")("coo_grade") = rs("grade")
		Response.Cookies("nkpmg_user")("coo_cost_grade") = rs("cost_grade")
		Response.Cookies("nkpmg_user")("coo_insa_grade") = rs("insa_grade")
		Response.Cookies("nkpmg_user")("coo_pay_grade") = rs("pay_grade")
		Response.Cookies("nkpmg_user")("coo_met_grade") = rs("met_grade")
		Response.Cookies("nkpmg_user")("coo_account_grade") = rs("account_grade")
		Response.Cookies("nkpmg_user")("coo_sales_grade") = rs("sales_grade")
		Response.Cookies("nkpmg_user")("coo_user_grade") = rs("user_grade")
		Response.Cookies("nkpmg_user")("coo_mg_group") = rs("mg_group")
		Response.Cookies("nkpmg_user")("coo_reside") = rs("reside")
		Response.Cookies("nkpmg_user")("coo_help_yn") = rs("help_yn")

		' email을 검사해서 @ 앞을것을 그룹웨어의 id 로 삼는다.
		If IsNull(rs("email")) Then
	  		Response.Cookies("nkpmg_user")("coo_groupware_id") = ""
	  	Else
		  If (rs("email")="") Then
			Response.Cookies("nkpmg_user")("coo_groupware_id") = ""
		  Else
			Response.Cookies("nkpmg_user")("coo_groupware_id") = split(rs("email"),"@")(0)
		  End If
		End If

		If IsNull(rs("position")) Then
			Response.Cookies("nkpmg_user")("coo_position") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_position") = rs("position")
		End If

		If IsNull(rs("org_name")) Then
			Response.Cookies("nkpmg_user")("coo_org_name") = ""
		Else
			Response.Cookies("nkpmg_user")("coo_org_name") = rs("org_name")
		End If

		If IsNull(rs("asset_company")) Then
			Response.Cookies("nkpmg_user")("coo_asset_company") = "00"
		Else
			Response.Cookies("nkpmg_user")("coo_asset_company") = rs("asset_company")
		End If

		login_cnt = int(rs("login_cnt")) + 1

		sql = "UPDATE memb SET login_cnt="&login_cnt&", login_date=now() WHERE user_id='"&id&"'"
		DBConn.Execute(sql)

		If (rs("mg_group") > "5" Or rs("grade") = "5" Or rs("team") = "외주관리") And IsNull(rs("asset_company")) Then
			Response.write "<script type='text/javascript'>"
			Response.write "	location.replace('as_list_ce_user.asp');"
			Response.write "</script>"
		'조건 중복으로 주석 처리[허정호_20201117]
		'ElseIf rs("grade") = "5" And isnull(rs("asset_company")) Then
		'	Response.write "<script type='text/javascript'>"
		'	Response.write "	location.replace('as_list_ce_user.asp');"
		'	Response.write "</script>"
		ElseIf rs("grade") = "5" And rs("asset_company") > "00" And rs("reside") = "2" Then
			Response.write "<script type='text/javascript'>"
			Response.write "	location.replace('asset_process_mg.asp');"
			Response.write "</script>"
		ElseIf rs("grade") = "5" And rs("asset_company") > "00" And rs("reside") = "3" Then
			Response.write "<script type='text/javascript'>"
			Response.write "	location.replace('as_list_asset.asp');"
			Response.write "</script>"
		ElseIf rs("grade") = "5" And rs("asset_company") > "00" And rs("reside") < "2" Then
			Response.write "<script language=javascript>"
			'response.write "location.replace('k1_asset_user_sum.asp');"
			Response.write "location.replace('as_list_asset.asp');"
			Response.write "</script>"
		Else
			Response.write "<script type='text/javascript'>"
			Response.write "	location.replace('nkp_main.asp?first_sw=y');"
			Response.write "</script>"
		End If
	End If

	'레코드 객체 제거 추가[허정호_20201117]
	rs.Close()
	Set rs = Nothing

	'DB 연결 객체 제거
	DBConn.Close()
	Set DBConn = Nothing

	Response.End
%>
