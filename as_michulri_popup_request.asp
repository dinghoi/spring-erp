<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim company_tab(50)

from_date = request("from_date")
to_date = request("to_date")
sido = request("sido")
mg_ce = request("mg_ce")
mg_ce_id = request("mg_ce_id")
mg_group = request("mg_group")
company = request("company")
as_type = request("as_type")
win_sw = "back"

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

i = 0
in_cnt = 0
acpt_cnt = 0
yun_cnt = 0

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

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

if company = "��ü" then
	grade_sql = ""
  else
	grade_sql = "( company = '" + company + "') and "
end if
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

com_sql = grade_sql

if	mg_ce = "" then
  if   sido = "��" then
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"')"
  elseif sido = "����" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','���','��õ')"
  elseif sido = "�λ�����" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('�λ�','�泲','���')"
  elseif sido = "�뱸����" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('�뱸','���')"
  elseif sido = "��������" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','�泲','���','����')"
  ' ������ ��������� ���� (2018.09.27 ����) 
  elseif sido = "��������" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����','����','����')"
  ' �������簡 ������ (2018.09.27 ����) 
  'elseif sido = "��������" then 
  '  sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
  '  sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
  elseif sido = "��������" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
  elseif sido = "��������" then 
    sql = "select * from as_acpt WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '�԰�' or as_process = '����')"
    sql = sql + " and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') and sido in ('����')"
  else
		sql = "select * from as_acpt"
		sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '����' or as_process = '�԰�') and (sido = '" + sido + "')"
		sql = sql + "  and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') Order By acpt_date Asc"
	end if
  else
	sql = "select * from as_acpt"
	sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '����' or as_process = '�԰�') and (mg_ce_id = '" + mg_ce_id + "')"
	sql = sql + "  and (Cast(request_date as date) >= '" + from_date + "' AND Cast(request_date as date) <= '"+to_date+"') Order By acpt_date Asc"
end if
if  from_date = "" then
	sql = "select * from as_acpt"
	sql = sql + " WHERE "+com_sql+type_sql+" (as_process = '����' or as_process = '����' or as_process = '�԰�') and (sido = '" + sido + "')"
	sql = sql + " Order By acpt_date Asc"
end if
Rs.Open Sql, Dbconn, 1

title_line = "��ó�� ��Ȳ"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��ó�� ��Ȳ</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}

        </script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="10%" >
							<col width="12%" >
							<col width="10%" >
							<col width="20%" >
							<col width="10%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th><%=memo01%></th>
							  <td class="left"><%=memo02%></td>
							  <th>ȸ��</th>
							  <td class="left"><%=company%></td>
							  <th>ó������</th>
							  <td class="left"><%=as_type%></td>
							  <td><a href = "as_michulri_excel_request.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&sido=<%=sido%>&company=<%=company%>&as_type=<%=as_type%>&mg_ce=<%=mg_ce%>&mg_ce_id=<%=mg_ce_id%>&mg_group=<%=mg_group%>" class="btnType04">�����ٿ�ε�</a>
							  </td>
					      	</tr>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="15%" >
							<col width="5%" >
							<col width="18%" >
							<col width="25%" >
							<col width="*" >
							<col width="10%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��û����</th>
								<th scope="col">����</th>
								<th scope="col">ȸ���</th>
								<th scope="col">�μ���</th>
								<th scope="col">����</th>
								<th scope="col">ó������</th>
								<th scope="col">��ȸ</th>
							</tr>
						</thead>
						<tbody>
					  	<%
                        do until rs.eof
                            int date_len 
                    '		date_len=len(rs("acpt_date"))
                            dim len_date, hangle, bit01, bit02, bit03
                            acpt_date = rs("acpt_date")
                            len_date = len(acpt_date)
                            bit01 = left(acpt_date, 10)
                        ' 	bit01 = Replace(bit01,"-",".")
                            bit03 = left(right(acpt_date, 5), 2)
                            hangle = mid(acpt_date, 12, 2)
                            if len_date = 22 then
                                bit02 = mid(acpt_date, 15, 2)
                              else
                                bit02 = "0"&mid(acpt_date, 15, 1)
                            end If
                         
                            if hangle = "����" and bit02 <> 12 then 
                                bit02 = bit02 + 12
                            end if
                            
                            date_to_date = bit01 & " " &bit02 & ":" & bit03
                            acpt_date = mid(date_to_date,3)
                            acpt_date = replace(acpt_date,"-","/")
                            acpt_date = rs("request_date")
                            if rs("as_process") = "����" then
                                acpt_cnt = acpt_cnt + 1
                            end if
                            if rs("as_process") = "����" then
                                yun_cnt = yun_cnt + 1
                            end if
                            if rs("as_process") = "�԰�" then
                                in_cnt = in_cnt + 1
                            end if
                    
                            i = i + 1
                      	%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=acpt_date%></td>
								<td><%=rs("as_process")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=rs("as_type")%></td>
								<td><a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=rs("acpt_no")%>&win_sw=<%=win_sw%>','asview_pop','scrollbars=yes,width=800,height=700')">��ȸ</a></td>
							</tr>
							<%
                                rs.movenext()
                            loop
                            %>
						</tbody>
					</table>                    
					<br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>����</th>
							  <td class="left"><%=acpt_cnt%></td>
							  <th>����</th>
							  <td class="left"><%=yun_cnt%></td>
							  <th>�԰�</th>
							  <td class="left"><%=in_cnt%></td>
					      	</tr>
						</tbody>
					</table>
					<br>
				</form>
				</div>
			</div>
	</body>
</html>

