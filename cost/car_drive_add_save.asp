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
Dim u_type, mg_ce_id, run_date, old_date, run_seq, old_run_seq, car_no, car_name
Dim oil_kind, car_owner, last_km, start_company, start_point, start_km, start_mm
Dim start_time, end_company, end_point, end_km, end_hh, end_mm, end_time
Dim run_memo, far, start_hh, oil_amt, oil_pay, oil_price, parking_pay, parking
Dim toll_pay, toll, cancel_yn, end_yn, mFormd_id, mod_user, mod_date, somopum
Dim mod_id, company, end_msg

u_type = Request.Form("u_type")
mg_ce_id = Request.Form("mg_ce_id")
run_date = Request.Form("run_date")
old_date = Request.Form("old_date")
run_seq = Request.Form("run_seq")
old_run_seq = Request.Form("run_seq")
car_no = Request.Form("car_no")
car_name = Request.Form("car_name")
oil_kind = Request.Form("oil_kind")
car_owner = Request.Form("car_owner")
last_km = Int(Request.Form("last_km"))
start_company = Request.Form("start_company")
start_point = Request.Form("start_point")
start_km = Int(Request.Form("start_km"))
start_hh = Request.Form("start_hh")
start_mm = Request.Form("start_mm")
end_company = Request.Form("end_company")
end_point = Request.Form("end_point")
end_km = Int(Request.Form("end_km"))
end_hh = Request.Form("end_hh")
end_mm = Request.Form("end_mm")
run_memo = Request.Form("run_memo")
far = Int(Request.Form("far"))
oil_amt = Request.Form("oil_amt")
oil_pay = Request.Form("oil_pay")
oil_price = Request.Form("oil_price")
parking_pay = Request.Form("parking_pay")
parking = Request.Form("parking")
toll_pay = Request.Form("toll_pay")
toll = Request.Form("toll")
cancel_yn = Request.Form("cancel_yn")
end_yn = Request.Form("end_yn")
mFormd_id = Request.Form("mod_id")
mod_user = Request.Form("mod_user")
mod_date = Request.Form("mod_date")

start_time = CStr(start_hh)&CStr(start_mm)
start_time = CStr(start_hh)&CStr(start_mm)
end_time = CStr(end_hh)&CStr(end_mm)

If car_owner = "개인" Then
	somopum = far * 25
Else
	somopum = 0
End If

If f_toString(oil_amt, "") = "" Then
	oil_amt = 0
End If
oil_amt = Int(oil_amt)

If f_toString(oil_price, "") = "" Then
	oil_price = 0
End If
oil_price = Int(oil_price)

If f_toString(parking, "") = "" Then
	parking = 0
End If
parking = Int(parking)

If f_toString(toll, "") = "" Then
	toll = 0
End If
toll = Int(toll)

If f_toString(mod_id, "") <> "" Then
	mod_yymmdd = DateValue(mod_date)
	mod_hhmm = FormatDateTime(mod_date,4)
	mod_date = CStr(mod_yymmdd)&" "&CStr(mod_hhmm)
End If

company = end_company

If company = "집" Or company = "본사(회사)" Or company = "기타" Or company = "케이원정보통신" Then
	company = start_company
End If

If company = "집" Or company = "본사(회사)" Or company = "기타" Or company = "케이원정보통신" Then
	company = "공통"
End If

If IsNull(reside_company) Then
	reside_company = ""
End If

If company = "공통" And reside_company <> "" Then
	company = reside_company
End If

DBConn.BeginTrans

If u_type = "U" Then
	'sql = "delete from transit_cost where run_date ='"&old_date&"' and mg_ce_id='"&mg_ce_id&"' and run_seq= '"&run_seq&"'"
	objBuilder.Append "DELETE FROM transit_cost "
	objBuilder.Append "WHERE run_date ='"&old_date&"' AND mg_ce_id='"&mg_ce_id&"' AND run_seq= '"&run_seq&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

Dim rsMaxSeq, max_seq

'sql = "select max(run_seq) as max_seq from transit_cost where mg_ce_id = '"&mg_ce_id&"' and run_date = '"&run_date&"'"
objBuilder.Append "SELECT MAX(run_seq) AS 'max_seq' FROM transit_cost "
objBuilder.Append "WHERE mg_ce_id = '"&mg_ce_id&"' AND run_date = '"&run_date&"';"

Set rsMaxSeq = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsMaxSeq.EOF Or rsMaxSeq.BOF Then
	max_seq = 0
Else
	max_seq = rsMaxSeq("max_seq")
End If

If IsNull(rsMaxSeq("max_seq")) Then
	max_seq = 0
End If

max_seq = max_seq + 1

If max_seq < 10 Then
	run_seq = "0"&CStr(max_seq)
Else
	run_seq = CStr(max_seq)
End If
rsMaxSeq.Close():Set rsMaxSeq = Nothing

If run_date = old_date Then
	run_seq = old_run_seq
End If

Dim rs_memb, emp_name, emp_grade

'sql = "select * from memb where user_id = '"&mg_ce_id&"'"
objBuilder.Append "SELECT user_name, user_grade FROM memb "
objBuilder.Append "WHERE user_id = '"&mg_ce_id&"' AND grade < '5';"

Set rs_memb = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

emp_name = rs_memb("user_name")
emp_grade = rs_memb("user_grade")

rs_memb.Close() : Set rs_memb = Nothing

'If f_toString(mod_id, "") = "" Then
	'sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_no,car_name,car_owner,oil_kind,start_company,start_point,start_km,start_time,end_company,end_point,end_km,end_time,far,run_memo,company,somopum"& _
	'",oil_amt,oil_pay,oil_price,parking_pay,parking,toll_pay,toll,cancel_yn,end_yn,reg_id,reg_user,reg_date) values ("& _
	'"'"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place& _
	'"','"&car_no&"','"&car_name&"','"&car_owner&"','"&oil_kind&"','"&start_company&"','"&start_point&"',"&start_km&",'"&start_time& _
	'"','"&end_company&"','"&end_point&"',"&end_km&",'"&end_time&"',"&far&",'"&run_memo&"','"&company&"',"&somopum& _
	'","&oil_amt&",'"&oil_pay&"',"&oil_price&",'"&parking_pay&"',"&parking&",'"&toll_pay&"',"&toll&",'"&cancel_yn&"','"&end_yn& _
	'"','"&user_id&"','"&user_name&"',now())"

	'DBConn.Execute(sql)
'Else
	'sql="insert into transit_cost (mg_ce_id,user_name,user_grade,run_date,run_seq,emp_company,bonbu,saupbu,team,org_name,reside_place,car_no,car_name,car_owner"& _
	'",oil_kind,start_company,start_point,start_km,start_time,end_company,end_point,end_km,end_time,far,run_memo,company,somopum"& _
	'",oil_amt,oil_pay,oil_price,parking_pay,parking,toll_pay,toll,cancel_yn,end_yn,reg_id,reg_user,reg_date,mod_id,mod_user"& _
	'",mod_date) values ('"&mg_ce_id&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"','"&run_date&"','"&run_seq&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name& _
	'"','"&reside_place&"','"&car_no&"','"&car_name&"','"&car_owner&"','"&oil_kind&"','"&start_company&"','"&start_point&"',"&start_km& _
	'",'"&start_time&"','"&end_company&"','"&end_point&"',"&end_km&",'"&end_time&"',"&far&",'"&run_memo&"','"&company&"',"&somopum& _
	'","&oil_amt&",'"&oil_pay&"',"&oil_price&",'"&parking_pay&"',"&parking&",'"&toll_pay&"',"&toll&",'"&cancel_yn& _
	'"','"&end_yn&"','"&user_id&"','"&user_name&"',now(),'"&mod_id&"','"&mod_user&"','"&mod_date&"')"

	'DBConn.Execute(sql)
'End If

objBuilder.Append "INSERT INTO transit_cost(mg_ce_id, user_name, user_grade, run_date, run_seq,"
objBuilder.Append "emp_company, bonbu, saupbu, team, org_name, reside_place,"
objBuilder.Append "car_no, car_name, car_owner, oil_kind, start_company,"
objBuilder.Append "start_point, start_km, start_time, end_company, end_point,"
objBuilder.Append "end_km, end_time, far, run_memo, company,"
objBuilder.Append "somopum, oil_amt, oil_pay, oil_price, parking_pay, parking,"
objBuilder.Append "toll_pay, toll, cancel_yn, end_yn, reg_id, reg_user, reg_date"

If f_toString(mod_id, "") <> "" Then
	objBuilder.Append ", mod_id, mod_user, mod_date"
End If

objBuilder.Append ")VALUES("
objBuilder.Append "'"&mg_ce_id&"','"&emp_name&"','"&emp_grade&"','"&run_date&"','"&run_seq&"',"
objBuilder.Append "'"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"',"
objBuilder.Append "'"&car_no&"','"&car_name&"','"&car_owner&"','"&oil_kind&"','"&start_company&"',"
objBuilder.Append "'"&start_point&"',"&start_km&",'"&start_time&"','"&end_company&"','"&end_point&"',"
objBuilder.Append ""&end_km&",'"&end_time&"',"&far&",'"&run_memo&"','"&company&"',"
objBuilder.Append ""&somopum&","&oil_amt&",'"&oil_pay&"',"&oil_price&",'"&parking_pay&"',"&parking&","
objBuilder.Append "'"&toll_pay&"',"&toll&",'"&cancel_yn&"','"&end_yn&"','"&user_id&"','"&user_name&"',NOW()"

If f_toString(mod_id, "") <> "" Then
	objBuilder.Append ",'"&mod_id&"','"&mod_user&"','"&mod_date&"'"
End If

objBuilder.Append ");"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If end_km > last_km Then
	'sql = "Update car_info set last_km="&end_km&" where car_no = '"&car_no&"'"
	objBuilder.Append "UPDATE car_info SET "
	objBuilder.Append "	last_km="&end_km
	objBuilder.Append "WHERE car_no = '"&car_no&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "변경 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "변경되었습니다."
End If

DBConn.Close():Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write"</script>"
Response.End
%>
