<!--#include virtual = "/common/inc_top.asp"--><!--���� ����-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<%
'==========================
'author : ����ȣ
'modify date : 20201117
'Desc :
'	���� ���� include �߰�
'	���� ���� �߰� �� ��� ��ü �Ҹ� ó��
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
		Response.write"alert('���̵� ��ϵǾ� ���� �ʽ��ϴ�....');"
		Response.write"history.go(-1);"
		Response.write"</script>"
	End If

	Set DBConn = Server.CreateObject("ADODB.CONNECTION")
	DBConn.Open DbConnect

    ' ���๮�� ���� �ڿ��� ��û 2019-06-04
    ' ���� ���� ��찡 �־� �ƿ� �α��ν� ��üó��.. as��Ͻ� �۾��� �߰����� ������ ��..
    sql="UPDATE memb SET org_name = REPLACE(org_name, '\r', '') WHERE org_name LIKE '%\r%' "
    DBConn.Execute(sql)

    sql="UPDATE memb SET org_name = REPLACE(org_name, '\n', '') WHERE org_name LIKE '%\n%' "
    DBConn.Execute(sql)
    ' ���๮�� ���� �ڿ��� ��û 2019-06-04

	'��ȸ �÷� ǥ��[����ȣ_20201117]
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
		Response.write "	alert('���̵� ��ϵǾ� ���� �ʽ��ϴ�....');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	ElseIf rs("pass") <> pass Then
		Response.write "<script language=javascript>"
		Response.write "	alert('��й�ȣ�� �ٸ��ϴ�....');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	ElseIf rs("grade") = "6" Then
		Response.write "<script language=javascript>"
		Response.write "	alert('NKP �������� ���ų� ������������ ����� �ٲ�� �ֽ��ϴ�.');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	ElseIf rs("grade") = "" Or IsNull(rs("grade")) Then
		Response.write "<script language=javascript>"
		Response.write "	alert('�α��� �Ҽ� �����ϴ�. �����ڿ��� ���� �ٶ��ϴ�....');"
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

		' email�� �˻��ؼ� @ �������� �׷������ id �� ��´�.
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

		If (rs("mg_group") > "5" Or rs("grade") = "5" Or rs("team") = "���ְ���") And IsNull(rs("asset_company")) Then
			Response.write "<script type='text/javascript'>"
			Response.write "	location.replace('as_list_ce_user.asp');"
			Response.write "</script>"
		'���� �ߺ����� �ּ� ó��[����ȣ_20201117]
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

	'���ڵ� ��ü ���� �߰�[����ȣ_20201117]
	rs.Close()
	Set rs = Nothing

	'DB ���� ��ü ����
	DBConn.Close()
	Set DBConn = Nothing

	Response.End
%>
