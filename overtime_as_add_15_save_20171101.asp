<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	on Error resume next

	acpt_no = request.form("acpt_no")
	work_item = request.form("work_item")
	work_date = request.form("work_date")
	company = request.form("company")
	dept = request.form("dept")
	from_hh = request.form("from_hh")
	from_mm = request.form("from_mm")	
	from_time = cstr(from_hh) + cstr(from_mm)
	to_hh = request.form("to_hh")
	to_mm = request.form("to_mm")	
	to_time = cstr(to_hh) + cstr(to_mm)
	work_gubun = request.form("work_gubun")
	work_memo = work_item
	sign_no = request.form("sign_no")	
	you_yn = request.form("you_yn")	
'	cost_detail = work_gubun
'	if work_gubun = "���Ͼ߱�" or work_gubun = "Ư�ٹ���" or work_gubun = "Ư������" or work_gubun = "Ư�پ߱�" then
'		cost_detail = "�߱�"
'	end if
'	if work_gubun = "�����Ͼ߱�" or work_gubun = "��Ư��12�������" or work_gubun = "��Ư��13����̻�" or work_gubun = "��Ư�پ߱�" or work_gubun = "��Ư��ö��" then
'		cost_detail = "���߱�"
'	end if
	
	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql = "select * from overtime_code where work_gubun = '"&work_gubun&"'"
	set rs_etc=dbconn.execute(sql)
	cost_detail = rs_etc("cost_detail")
	overtime_amt = rs_etc("overtime_amt")
	
	sql = "select * from ce_work where work_id = '2' and acpt_no ="&int(acpt_no)
	Rs.Open Sql, Dbconn, 1

	do until rs.eof
		sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
		set rs_memb=dbconn.execute(sql)		

		sql="insert into overtime (work_date,mg_ce_id,user_name,user_grade,acpt_no,emp_company,bonbu,saupbu,team,org_name,reside_place,company,dept,work_item,from_time,to_time,work_gubun,cost_detail,person_amt,overtime_amt,work_memo,sign_no,you_yn,cancel_yn,end_yn,reg_id,reg_user,reg_date) values ('"&rs("work_date")&"','"&rs("mg_ce_id")&"','"&rs_memb("user_name")&"','"&rs_memb("user_grade")&"',"&rs("acpt_no")&",'"&rs("emp_company")&"','"&rs("bonbu")&"','"&rs("saupbu")&"','"&rs("team")&"','"&rs("org_name")&"','"&rs("reside_place")&"','"&rs("company")&"','"&dept&"','"&work_item&"','"&from_time&"','"&to_time&"','"&work_gubun&"','"&cost_detail&"',"&rs("person_amt")&","&overtime_amt&",'"&work_memo&"','"&sign_no&"','"&you_yn&"','N','N','"&user_id&"','"&user_name&"',now())"
		dbconn.execute(sql)
		rs.movenext()
	loop                                       		

	sql = "Update as_acpt set overtime ='Y' where acpt_no ="&int(acpt_no)
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "�̹� ��ϵǾ� �־� ����� Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = "��ϵǾ����ϴ�...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	

%>
