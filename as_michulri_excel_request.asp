<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

	dim title_name
	dim company_tab(50)

	from_date = request("from_date")
	to_date = request("to_date")
	curr_date = datevalue(mid(cstr(now()),1,10))
	sido = request("sido")
	mg_ce = request("mg_ce")
	mg_ce_id = request("mg_ce_id")
	mg_group = request("mg_group")
	company = request("company")
	as_type = request("as_type")
	days = int(request("days"))
	
	if company = "" then
		company = "��ü"
		as_type = "��ü"
	end if
	
	if mg_ce = "" then
		memo01 = "�õ�"
		memo02 = sido
	  else
		memo01 = "�����"
		memo02 = mg_ce
	end if
	
	if as_type = "��ü" then
		type_sql = ""
	  else
		type_sql = " (as_type ='"+as_type+"') and "
	end if

	title_name = array("��������","������","�����","��ȭ��ȣ","ȸ��","������","�ּ�","CE��","��ֳ���","��û��","��û�ð�","ó�����","����","�԰����")

	if mg_ce = "" then
		title_memo = sido + " ������ "
	  else
	    title_memo = mg_ce + " ����� "
	end if
	savefilename = title_memo + "��ó�� ����.xls"

 	Response.Buffer = True
  	Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
  	Response.CacheControl = "public"
  	Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_hol = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Dbconn.open DbConnect

	if company = "��ü" and c_grade = "7" then
		k = 0
		Sql="select * from etc_code where etc_type = '51' and used_sw = 'Y' and group_name = '"+user_name+"' order by etc_name asc"
		Rs_etc.Open Sql, Dbconn, 1
		while not rs_etc.eof
			k = k + 1
			company_tab(k) = rs_etc("etc_name")
			rs_etc.movenext()
		Wend
	rs_etc.close()						
	end if				
	
	grade_sql = "( company = '" + company + "') and "
	if c_grade = "7"  and company = "��ü" then
		com_sql = "company = '" + company_tab(1) + "'"	
		for kk = 2 to k
			com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
		next
		grade_sql = "(" + com_sql + ") and "
	end if
	
	if ( c_grade = "8" ) or (c_grade = "7"  and company <> "��ü") then
		grade_sql = "( company = '" + company + "') and "
	end if
	if c_grade <> "7" and company = "��ü" then
		grade_sql = " "
	end if
	
	com_sql = grade_sql

	' ��ó����
	if	mg_ce = "" then
		if sido = "�Ѱ�" then
			sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
			sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
			sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
			sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
		 
	  elseif   sido = "��" then
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
	  elseif sido = "����" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','���','��õ')"
	  elseif sido = "�λ�����" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('�λ�','�泲','���')"
	  elseif sido = "�뱸����" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('�뱸','���')"
	  elseif sido = "��������" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','�泲','���','����')"
	  elseif sido = "��������" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','����')"
	  elseif sido = "��������" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
	  elseif sido = "��������" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
	  elseif sido = "��������" then 
      sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
      sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
      sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
	    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
	  else
			sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
			sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
			sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
			sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and (sido = '" + sido + "')"
		end if	  
	  else
		if mg_ce = "�Ѱ�" then
			sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
			sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
			sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
			sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
		  else
			sql = "select acpt_date,acpt_man,acpt_user,replace(tel_ddd,' ','')+'-'+replace(tel_no1,' ','')+'-'+replace(tel_no2,' ',''),company,dept,sido+' '+gugun+' '+dong+' '+addr,"
			sql = sql + "mg_ce,as_memo,request_date,request_time,as_type,as_process,into_reason from as_acpt "
			sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
			sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and (mg_ce_id = '" + mg_ce_id + "')"
		end if
	end if

	Rs.Open Sql, Dbconn, 1
	if rs.eof then
		response.write"<script language=javascript>"
		response.write"alert('�ٿ� �� �ڷᰡ �����ϴ� ....');"
		response.write"history.go(-1);"
		response.write"</script>"	
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title></title>
</head>
<body>
<table border='1' cellspacing='0' cellpadding='5' bordercolordark='white' bordercolorlight='black'>
	<tr><%=chr(13)&chr(10)%>
<%	
	i = 0
	for each whatever in rs.fields	
		if i < 14 then
%>
			<td><b><%=title_name(i)%></b></TD><%=chr(13)&chr(10)%>
<%		
		end if
		i = i + 1
	next	%>
	</tr><%=chr(13)&chr(10)%>
<%
alldata=rs.getrows

numcols=ubound(alldata,1)
numrows=ubound(alldata,2)

FOR j= 0 TO numrows 
%>
	<tr><%=chr(13)&chr(10)%>
<%  FOR i=0 to numcols
	if i > 13 then
		exit for
	end if
    thisfield=alldata(i,j)
      if isnull(thisfield) then
         thisfield=""
      end if
      if trim(thisfield)="" then
         thisfield=""
      end if
%>
<%	if i = 0 then %>
		<td valign=top><%=thisfield%> </td><%=chr(13)&chr(10)%>
<%		else	%>
		<td style="mso-number-format:'\@'" valign=top><%=thisfield%> </td><%=chr(13)&chr(10)%>
<%	end if 		%>
<%  NEXT	%>
	</tr><%=chr(13)&chr(10)%>
<%NEXT%>
</table>

</body>
</html>
