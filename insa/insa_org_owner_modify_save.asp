<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim u_type, org_level, org_code, org_date, org_empno
Dim org_empname, org_company, org_bonbu, org_saupbu, org_team
Dim org_cost_group, org_cost_center, org_reside_company, owner_org
Dim owner_orgname, owner_empno, owner_empname, org_table_org, org_zip
Dim org_sido, org_gugun, org_dong, org_addr, org_end_date, tel_ddd
Dim tel_no1, tel_no2, org_reside_place, org_reg_date, org_mod_date, org_owner_date
Dim end_msg

u_type = Request.Form("u_type")
org_level = Request.Form("org_level")
org_code = Request.Form("org_code")
org_name = Request.Form("org_name")
org_date = Request.Form("org_date")
org_empno = Request.Form("org_empno")
org_empname = Request.Form("org_empname")
org_company = Request.Form("org_company")
org_bonbu = Request.Form("org_bonbu")
org_saupbu = Request.Form("org_saupbu")
org_team = Request.Form("org_team")
org_cost_group = Request.Form("org_cost_group")
org_cost_center = Request.Form("org_cost_center")

If f_toString(org_bonbu, "") = "" Then
	   org_bonbu = ""
End If

If f_toString(org_saupbu, "") = "" Then
	   org_saupbu = ""
End If

If f_toString(org_team, "") = "" Then
	   org_team = ""
End If

'If org_level = "회사" Then
'	  org_company = org_name
'   elseif org_level = "본부" then
'			  org_bonbu = org_name
'		  elseif org_level = "사업부" then
'					 org_saupbu = org_name
'				 elseif org_level = "팀" then
'						   org_team = org_name
'end If

Select Case org_level
	Case "회사"
		org_company = org_name
	Case "본부"
		org_bonbu = org_name
	Case "사업부"
		org_saupbu = org_name
	Case "팀"
		org_team = org_name
End Select

org_reside_company = Request.Form("org_reside_company")

If f_toString(org_reside_company, "") = "" Then
	org_reside_company = ""
End If

If f_toString(org_cost_group, "") = "" Then
	   org_cost_group = org_reside_company
End If

owner_org = Request.Form("owner_org")
owner_orgname = Request.Form("owner_orgname")
owner_empno = Request.Form("owner_empno")
owner_empname = Request.Form("owner_empname")
org_table_org = Int(Request.Form("org_table_org"))
org_zip = Request.Form("org_zip")
org_sido = Request.Form("org_sido")
org_gugun = Request.Form("org_gugun")
org_dong = Request.Form("org_dong")
org_addr = Request.Form("org_addr")
org_end_date = Request.Form("org_end_date")
tel_ddd = Request.Form("tel_ddd")
tel_no1 = Request.Form("tel_no1")
tel_no2 = Request.Form("tel_no2")

If f_toString(tel_ddd, "") = "" Then
   tel_ddd = ""
   tel_no1 = ""
   tel_no2 = ""
End If

org_reside_place = Request.Form("org_reside_place")

If org_level = "상주처" Then
	org_cost_center = "상주직접비"
Else
	org_cost_group = org_saupbu
	'org_reside_company = ""
	If org_saupbu = "" Then
		If org_bonbu = ""  Then
		  org_cost_group = org_company
		Else
		  org_cost_group = org_bonbu
		End If
	End If
End If

If org_reside_company <> "" Then
   org_cost_group = Request.Form("org_cost_group")
End If

If f_toString(org_date, "") = "" Then
	org_date = "0000-00-00"
End If

If f_toString(org_end_date, "") = "" Then
	org_end_date = "0000-00-00"
End If

org_reg_date = Request.Form("org_reg_date")
org_mod_date = Request.Form("org_mod_date")

org_owner_date = Request.Form("org_owner_date")

'reg_user = request.cookies("nkpmg_user")("coo_user_name")
'mod_user = request.cookies("nkpmg_user")("coo_user_name")

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_org_mst SET "
	objBuilder.Append "	org_level='"&org_level&"',org_company='"&org_company&"',org_bonbu='"&org_bonbu&"',"
	objBuilder.Append "	org_saupbu='"&org_saupbu&"',org_team='"&org_team&"',org_name='"&org_name&"',"
	objBuilder.Append "	org_reside_place='"&org_reside_place&"',org_reside_company='"&org_reside_company&"',org_cost_group='"&org_cost_group&"',"
	objBuilder.Append "	org_empno='"&org_empno&"',org_emp_name='"&org_empname&"',org_date='"&org_date&"',"
	objBuilder.Append "	org_tel_ddd='"&tel_ddd&"',org_tel_no1='"&tel_no1&"',org_tel_no2='"&tel_no2&"',"
	objBuilder.Append "	org_owner_org='"&owner_org&"',org_owner_empno='"&owner_empno&"',org_owner_empname='"&owner_empname&"',"
	objBuilder.Append "	org_table_org='"&org_table_org&"',org_sido='"&org_sido&"',org_gugun='"&org_gugun&"',"
	objBuilder.Append "	org_dong='"&org_dong&"',org_addr='"&org_addr&"',org_cost_center='"&org_cost_center&"',"
	objBuilder.Append "	org_end_date='"&org_end_date&"',org_owner_date='"&org_owner_date&"',org_mod_date=NOW(),org_mod_user='"&user_name&"' "
	objBuilder.Append "WHERE org_code = '"&org_code&"'"
Else
	objBuilder.Append "insert into emp_org_mst (org_code,org_level,org_company,org_bonbu,org_saupbu,"
	objBuilder.Append "org_team,org_name,org_reside_place,org_reside_company,org_cost_group,"
	objBuilder.Append "org_empno,org_emp_name,org_date,org_tel_ddd,org_tel_no1,"
	objBuilder.Append "org_tel_no2,org_cost_center,org_owner_org,org_owner_empno,org_owner_empname,"
	objBuilder.Append "org_table_org,org_sido,org_gugun,org_dong,org_addr,org_reg_date,org_reg_user)"
	objBuilder.Append "VALUES('"&org_code&"','"&org_level&"','"&org_company&"','"&org_bonbu&"','"&org_saupbu&"',"
	objBuilder.Append "'"&org_team&"','"&org_name&"','"&org_reside_place&"','"&org_reside_company&"','"&org_cost_group&"',"
	objBuilder.Append "'"&org_empno&"','"&org_empname&"','"&org_date&"','"&tel_ddd&"','"&tel_no1&"',"
	objBuilder.Append "'"&tel_no2&"','"&org_cost_center&"','"&owner_org&"','"&owner_empno&"','"&owner_empname&"',"
	objBuilder.Append "'"&org_table_org&"','"&org_sido&"','"&org_gugun&"','"&org_dong&"','"&org_addr&"',NOW(),'"&user_name&"')"
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "상위조직이 정상적으로 변경되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write	"<script type='text/javascript'>"
Response.Write	"	alert('"&end_msg&"');"
'Response.Write	"	location.replace('insa_org.asp');"
Response.Write	"	self.opener.location.reload();"
Response.Write	"	window.close();"
Response.Write	"</script>"
Response.End
%>
