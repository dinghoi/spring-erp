<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim u_type, mg_ce_id, run_date, old_date, oil_kind, run_seq
Dim start_point, start_hh, start_mm, company, end_point
Dim end_hh, end_mm, transit, payment, run_memo, fare, end_yn
Dim cancel_yn, mod_id, mod_user, mod_date, car_owner, start_time
Dim end_time, mod_yymmdd, mod_hhmm, rsTran, max_seq, rs_memb, modYn
Dim m_user_name, m_user_grade, end_msg

u_type = Request.Form("u_type")
mg_ce_id = Request.Form("mg_ce_id")
run_date = Request.Form("run_date")
old_date = Request.Form("old_date")
run_seq = Request.Form("run_seq")
oil_kind = Request.Form("oil_kind")
start_point = Request.Form("start_point")
start_hh = Request.Form("start_hh")
start_mm = Request.Form("start_mm")
company = Request.Form("company")
end_point = Request.Form("end_point")
end_hh = Request.Form("end_hh")
end_mm = Request.Form("end_mm")
transit = Request.Form("transit")
payment = Request.Form("payment")
run_memo = Request.Form("run_memo")
fare = Int(Request.Form("fare"))
end_yn = Request.Form("end_yn")
cancel_yn = Request.Form("cancel_yn")
mod_id = Request.Form("mod_id")
mod_user = Request.Form("mod_user")
mod_date = Request.Form("mod_date")

car_owner = "대중교통"
start_time = CStr(start_hh)&CStr(start_mm)
end_time = CStr(end_hh)&CStr(end_mm)

If mod_id <> "" Then
	mod_yymmdd = DateValue(mod_date)
	mod_hhmm = FormatDateTime(mod_date,4)
	mod_date = CStr(mod_yymmdd)&" "&CStr(mod_hhmm)
End If

DBConn.BeginTrans

If u_type = "U" Then
	'sql = "delete from transit_cost where run_date ='"&old_date&"' and mg_ce_id='"&mg_ce_id&"' and run_seq='"&run_seq&"'"
	objBuilder.Append "DELETE FROM transit_cost "
	objBuilder.Append "WHERE run_date='"&old_date&"' AND mg_ce_id='"&mg_ce_id&"' AND run_seq='"&run_seq&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

'sql = "select max(run_seq) as max_seq from transit_cost where mg_ce_id = '"&mg_ce_id&"' and run_date = '"&run_date&"'"
objBuilder.Append "SELECT MAX(run_seq) AS max_seq FROM transit_cost "
objBuilder.Append "WHERE mg_ce_id='"&mg_ce_id&"' AND run_date='"&run_date&"';"

Set rsTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsTran.EOF or rsTran.BOF then
	max_seq = 0
Else
	max_seq = Int(rsTran("max_seq"))
End If

If f_toString(rsTran("max_seq"), "") = "" Then
	max_seq = 0
End If

max_seq = max_seq + 1

If max_seq < 10 Then
	run_seq = "0"&CStr(max_seq)
Else
	run_seq = CStr(max_seq)
End If
rsTran.Close() : Set rsTran = Nothing

'sql = "select * from memb where user_id = '"&mg_ce_id&"'"
objBuilder.Append "SELECT user_name, user_grade FROM memb "
objBuilder.Append "WHERE user_id='"&mg_ce_id&"' AND grade < '5';"

Set rs_memb = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

m_user_name = rs_memb("user_name")
m_user_grade = rs_memb("user_grade")

rs_memb.Close() : Set rs_memb = Nothing

'수정 여부
If f_toString(mod_id, "") = "" Then
	modYn = "N"
Else
	modYn = "N"
End If

objBuilder.Append "INSERT INTO transit_cost(mg_ce_id,user_name,user_grade,run_date,run_seq,"
objBuilder.Append "emp_company,bonbu,saupbu,team,org_name,"
objBuilder.Append "reside_place,car_owner,start_point,start_km,start_time,"
objBuilder.Append "end_point,end_km,end_time,transit,payment,"
objBuilder.Append "fare,run_memo,company,cancel_yn,end_yn,"
objBuilder.Append "reg_id,reg_user,reg_date"

If modYn = "Y" Then
	objBuilder.Append ", mod_id, mod_user, mod_date"
End If

objBuilder.Append ")VALUES('"&mg_ce_id&"','"&m_user_name&"','"&m_user_grade&"','"&run_date&"','"&run_seq&"',"
objBuilder.Append "'"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"',"
objBuilder.Append "'"&reside_place&"','"&car_owner&"','"&start_point&"',0,'"&start_time&"',"
objBuilder.Append "'"&end_point&"',0,'"&end_time&"','"&transit&"','"&payment&"',"
objBuilder.Append ""&fare&",'"&run_memo&"','"&company&"','"&cancel_yn&"','"&end_yn&"',"
objBuilder.Append "'"&user_id&"','"&user_name&"',NOW()"

If modYn = "Y" Then
	objBuilder.Append ",'"&mod_id&"','"&mod_name&"','"&mod_date&"'"
End If

objBuilder.Append ");"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()


'if isnull(mod_id) or mod_id = "" then
'	sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_owner,start_point"& _
'	",start_km,start_time,end_point,end_km,end_time,transit,payment,fare,run_memo,company,cancel_yn,end_yn,reg_id,reg_user,reg_date) "& _
'	"values ('"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name& _
'	"','"&reside_place&"','"&car_owner&"','"&start_point&"',0,'"&start_time&"','"&end_point&"',0,'"&end_time&"','"&transit&"','"&payment& _
'	"',"&fare&",'"&run_memo&"','"&company&"','"&cancel_yn&"','"&end_yn&"','"&user_id&"','"&user_name&"',now())"
'	dbconn.execute(sql)
'else
'	sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_owner,start_point"& _
'	",start_km,start_time,end_point,end_km,end_time,transit,payment,fare,run_memo,company,cancel_yn,end_yn,reg_id,reg_user,reg_date,mod_id"& _
'	",mod_user,mod_date) values ('"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team& _
'	"','"&org_name&"','"&reside_place&"','"&car_owner&"','"&start_point&"',0,'"&start_time&"','"&end_point&"',0,'"&end_time&"','"&transit& _
'	"','"&payment&"',"&fare&",'"&run_memo&"','"&company&"','"&cancel_yn&"','"&end_yn&"','"&user_id&"','"&user_name&"',now(),'"&mod_id& _
'	"','"&mod_name&"','"&mod_date&"')"
'	dbconn.execute(sql)
'end if

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "저장 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "저장 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End
%>
