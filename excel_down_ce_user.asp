<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

	dim title_name
	dim company_tab(50)

	title_name = array("������ȣ","��������","������","����","�����","��ȭ��ȣ","�ڵ���","ȸ��","������","�ּ�","CE��","��ֳ���","��û��","��û�ð�","ó�����","����","����û","������","������","�𵨸�","�԰����","�԰����")

	savefilename = user_id + ".xls"
 	Response.Buffer = True
  	Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
  	Response.CacheControl = "public"
  	Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

	if reside = "9" then
		k = 0
		Sql="select * from trade where use_sw = 'Y' and group_name = '"+user_name+"' order by trade_name asc"
		rs_trade.Open Sql, Dbconn, 1
		do until rs_trade.eof
			k = k + 1
			company_tab(k) = rs_trade("trade_name")
			rs_trade.movenext()
		loop
		rs_trade.close()						
	end if
	
	if reside = "9" then
		com_sql = "company = '" + company_tab(1) + "'"	
		for kk = 2 to k
			com_sql = com_sql + " or company = '" + company_tab(kk) + "'"
		next
		condi_sql = " or " + com_sql + ") "
	  else
		condi_sql = " or company = '" + reside_company + "' or company = '" + user_name + "') "
	end if

	'//2017-06-07 ����Ƽǻó(���:900002) �α��ν� �������� ��� �˻��ϰ� ����
	If  user_id = "900002" Then
		condi_sql = " or company in ('������ǰ','������ũ��','�ڿ���') " & condi_sql
	End If
	
	order_Sql = " ORDER BY acpt_date desc"
	
'	where_sql = " WHERE (acpt_man = '" + user_name + "' or company = '" + reside_company + "' or company = '" + user_name + "') and "
	where_sql = " WHERE (acpt_man = '" + user_name + "'" + condi_sql
	base_sql = " and (as_process = '����' or as_process = '�԰�' or as_process = '����' or as_process = '��ü�԰�') "
		
	sql = "select acpt_no,acpt_date,acpt_man,acpt_grade,acpt_user,concat(tel_ddd,'-',tel_no1,'-',tel_no2),concat(hp_ddd,'-',hp_no1,'-',hp_no2),company,dept,concat(sido,' ',gugun,' ',dong,' ',addr),mg_ce,as_memo,request_date,request_time,as_type,as_process,visit_request_yn,maker,as_device,model_no,into_reason from as_acpt "
	sql = sql + where_sql + base_sql + order_sql
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
	for i = 0 to 21
'	for each whatever in rs.fields	
'		if i < 21 then
%>
			<td><b><%=title_name(i)%></b></TD><%=chr(13)&chr(10)%>
<%		
	next
'		end if
'		i = i + 1
'	next	
%>
	</tr><%=chr(13)&chr(10)%>
<%
alldata=rs.getrows

numcols=ubound(alldata,1) + 1
numrows=ubound(alldata,2)

FOR j= 0 TO numrows 
	in_process = ""
	if alldata(15,j) = "�԰�" then
		sql = "select into_date,in_process,in_place from as_into where acpt_no="&alldata(0,j)&" and in_seq="&"(select max(in_seq) from as_into where acpt_no="&alldata(0,j)&")"
		Set Rs_in=dbconn.execute(sql)
		if	Rs_in.eof then
				in_process = "����"
			else
				in_process = rs_in("in_process")
		end if
	end if
	if alldata(16,j) = "Y" then
		alldata(16,j) = "�湮��û"
	  else
		alldata(16,j) = ""
	end if	  	

%>
	<tr><%=chr(13)&chr(10)%>
<%  FOR i=0 to numcols
	if i = 21 then
    	thisfield = in_process 
	  else
		thisfield=alldata(i,j)
	end if
      if isnull(thisfield) then
         thisfield=""
      end if
      if trim(thisfield)="" then
         thisfield=""
      end if
%>
<%	if i = 1 or i = 11 then %>
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
