<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_cost = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if cost_grade = "0" then
	sql = "select * from emp_org_mst where (org_level = '�����') and org_name <> '�Ѱ���ǥ' group by org_name Order By org_name Asc"
  else
	sql = "select * from emp_org_mst where org_level = '�����' and org_name <> '�Ѱ���ǥ' and (org_name = '"&saupbu&"' or org_empno ='"&emp_no&"') group by org_name"
end if
Rs.Open Sql, Dbconn, 1

title_line = "��� ���� ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
			function frmcheck () {
					document.frm.submit();
			}
			
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_end_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�ֽ������� �ٽ� ��ȸ�ϱ�&nbsp;</strong>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="13%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">�����</th>
								<th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� �� �� �� �� Ȳ</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">���ο� ���� ó��</th>
								<th rowspan="2" scope="col">�����ڷ�</th>
								<th rowspan="2" scope="col">�����庸��</th>
								<th rowspan="2" scope="col">CEO����</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">�������</th>
							  <th scope="col">��������</th>
							  <th scope="col">������</th>
							  <th scope="col">ó������</th>
							  <th scope="col">�������</th>
							  <th scope="col">�������</th>
							  <th scope="col">����ó��</th>
                          </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

							cancel_yn = "N"
							if rs("org_bonbu") = "���һ����" then
								if rs("org_saupbu") = "�������������" or rs("org_saupbu") = "KAL���������" then
									jik_yn = "N"
								  else
									jik_yn = "Y"
							  	end if
							  else
							  	jik_yn = "N"
							end if
							
							sql="select max(end_month) as max_month from cost_end where saupbu='"&rs("org_name")&"'"
							set rs_max=dbconn.execute(sql)

							sql="select * from cost_end where saupbu='"&rs("org_name")&"' and end_month ='"&rs_max("max_month")&"'"
							set rs_cost=dbconn.execute(sql)
							if rs_cost.eof or rs_cost.bof then
								new_date = dateadd("m",-1,now())
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
								end_month = "����"
								end_view = ""
								batch_yn = ""
								bonbu_yn = ""
								ceo_yn = ""
								batch_view = ""
								ceo_view = ""
								reg_name = ""
								reg_id = ""
								reg_date = ""
							  else														
								new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
									
								if  rs_cost("end_yn") = "Y" then
									end_view = "����"
								  elseif rs_cost("end_yn") = "C" then
									new_month = rs_cost("end_month")
									end_view = "���"
								  else
									end_view = "����"
								end if
								end_yn = rs_cost("end_yn")
								end_month = rs_cost("end_month")							
								reg_name = rs_cost("reg_name")
								reg_id = rs_cost("reg_id")
								reg_date = rs_cost("reg_date")
								if rs_cost("batch_yn") = "Y" then
									batch_view = "�ڷ����"
								  else
								  	batch_view = "�̻���"
								end if
								if rs_cost("bonbu_yn") = "Y" then
									bonbu_view = "���οϷ�"
								end if
								if rs_cost("ceo_yn") = "Y" then
									ceo_view = "���οϷ�"
								end if
								if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
									bonbu_view = "������"
								  	ceo_view = ""
								end if								
								if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
								  	ceo_view = "������"
								end if								
								if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
									bonbu_view = ""
								  	ceo_view = ""
								end if								
								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
							end if					
							if jik_yn = "Y" then
								if ceo_yn = "N" then
									cancel_yn = "Y"
								end if
							  else
							  	if bonbu_yn = "N" then
									cancel_yn = "Y"
								end if
							end if									
						%>
							<tr>
								<td class="first"><%=rs("org_name")%></td>
								<td><%=end_month%></td>
								<td><%=end_view%>&nbsp;</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
							<% if cancel_yn = "Y" then	%>
                                <a href="cost_end_cancel.asp?saupbu=<%=rs("org_name")%>&end_month=<%=end_month%>" class="btnType03">�������</a>
							<%   else	%>
								��ҺҰ�
                            <% end if	%>
                                </td>
								<td><input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true"></td>
								<td>
							<% if now_month > new_month then	%>
                                <a href="cost_end_pro.asp?saupbu=<%=rs("org_name")%>&end_month=<%=new_month%>&end_yn=<%=end_yn%>" class="btnType03">����</a>
							<%   else	%>
								�����Ұ�
                            <% end if	%>
                                </td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
				    <%
							sql="select max(end_month) as max_month from cost_end where saupbu='����οܳ�����'"
							set rs_max=dbconn.execute(sql)

							sql="select * from cost_end where saupbu='����οܳ�����' and end_month ='"&rs_max("max_month")&"'"
							set rs_cost=dbconn.execute(sql)
							if rs_cost.eof or rs_cost.bof then
								new_date = "2015-01-01"
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
								end_month = "����"
								end_view = ""
								batch_yn = ""
								bonbu_yn = ""
								ceo_yn = ""
								batch_view = ""
								ceo_view = ""
								reg_name = ""
								reg_id = ""
								reg_date = ""
							  else														
								new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
									
								if  rs_cost("end_yn") = "Y" then
									end_view = "����"
								  elseif rs_cost("end_yn") = "C" then
									new_month = rs_cost("end_month")
									end_view = "���"
								  else
									end_view = "����"
								end if
								end_yn = rs_cost("end_yn")
								end_month = rs_cost("end_month")							
								reg_name = rs_cost("reg_name")
								reg_id = rs_cost("reg_id")
								reg_date = rs_cost("reg_date")
								if rs_cost("batch_yn") = "Y" then
									batch_view = "�ڷ����"
								  else
								  	batch_view = "�̻���"
								end if
								if rs_cost("bonbu_yn") = "Y" then
									bonbu_view = "���οϷ�"
								end if
								if rs_cost("ceo_yn") = "Y" then
									ceo_view = "���οϷ�"
								end if
								if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
									bonbu_view = "������"
								  	ceo_view = ""
								end if								
								if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
								  	ceo_view = "������"
								end if								
								if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
									bonbu_view = ""
								  	ceo_view = ""
								end if								
								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
							end if					
							if jik_yn = "Y" then
								if ceo_yn = "N" then
									cancel_yn = "Y"
								end if
							  else
							  	if bonbu_yn = "N" then
									cancel_yn = "Y"
								end if
							end if									
						%>

							<tr bgcolor="#FFE8E8">
								<td class="first">����οܳ�����</td>
								<td><%=end_month%></td>
								<td><%=end_view%>&nbsp;</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
							<% if cancel_yn = "Y" then	%>
                                <a href="cost_bonbu_end_cancel.asp?saupbu=<%="����οܳ�����"%>&end_month=<%=end_month%>" class="btnType03">�������</a>
							<%   else	%>
								��ҺҰ�
                            <% end if	%>
                                </td>
								<td><input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true"></td>
								<td>
							<% if now_month > new_month then	%>
                                <a href="cost_bonbu_end_pro.asp?saupbu=<%="����οܳ�����"%>&end_month=<%=new_month%>&end_yn=<%=end_yn%>" class="btnType03">����</a>
							<%   else	%>
								�����Ұ�
                            <% end if	%>
                                </td>
								<td><%=batch_view%>&nbsp;</td>
								<td><%=bonbu_view%>&nbsp;</td>
								<td><%=ceo_view%>&nbsp;</td>
							</tr>
						<%
							sql="select max(end_month) as max_month from cost_end where saupbu='���ֺ��'"
							set rs_max=dbconn.execute(sql)

							sql="select * from cost_end where saupbu='���ֺ��' and end_month ='"&rs_max("max_month")&"'"
							set rs_cost=dbconn.execute(sql)
							if rs_cost.eof or rs_cost.bof then
								new_date = "2015-01-01"
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
								end_month = "����"
								end_view = ""
								batch_yn = ""
								bonbu_yn = ""
								ceo_yn = ""
								batch_view = ""
								ceo_view = ""
								reg_name = ""
								reg_id = ""
								reg_date = ""
							  else														
								new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
									
								if  rs_cost("end_yn") = "Y" then
									end_view = "����"
								  elseif rs_cost("end_yn") = "C" then
									new_month = rs_cost("end_month")
									end_view = "���"
								  else
									end_view = "����"
								end if
								end_yn = rs_cost("end_yn")
								end_month = rs_cost("end_month")							
								reg_name = rs_cost("reg_name")
								reg_id = rs_cost("reg_id")
								reg_date = rs_cost("reg_date")
								if rs_cost("batch_yn") = "Y" then
									batch_view = "�ڷ����"
								  else
								  	batch_view = "�̻���"
								end if
								if rs_cost("bonbu_yn") = "Y" then
									bonbu_view = "���οϷ�"
								end if
								if rs_cost("ceo_yn") = "Y" then
									ceo_view = "���οϷ�"
								end if
								if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
									bonbu_view = "������"
								  	ceo_view = ""
								end if								
								if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
								  	ceo_view = "������"
								end if								
								if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
									bonbu_view = ""
								  	ceo_view = ""
								end if								
								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
							end if					
							if jik_yn = "Y" then
								if ceo_yn = "N" then
									cancel_yn = "Y"
								end if
							  else
							  	if bonbu_yn = "N" then
									cancel_yn = "Y"
								end if
							end if									
						%>

							<tr bgcolor="#FFFFCC">
								<td class="first">���ֺ��</td>
								<td><%=end_month%></td>
								<td><%=end_view%>&nbsp;</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
							<% if cancel_yn = "Y" then	%>
                                <a href="company_cost_end_cancel.asp?end_month=<%=end_month%>" class="btnType03">�������</a>
							<%   else	%>
								��ҺҰ�
                            <% end if	%>
                                </td>
								<td><input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true"></td>
								<td>
							<% if now_month > new_month then	%>
                                <a href="company_cost_end_pro.asp?end_month=<%=new_month%>&end_yn=<%=end_yn%>" class="btnType03">����</a>
							<%   else	%>
								�����Ұ�
                            <% end if	%>
                           	  </td> 
							  <td colspan="3">&nbsp;</td>
							</tr>
						<%
							sql="select max(end_month) as max_month from cost_end where saupbu='�κа������'"
							set rs_max=dbconn.execute(sql)

							sql="select * from cost_end where saupbu='�κа������' and end_month ='"&rs_max("max_month")&"'"
							set rs_cost=dbconn.execute(sql)
							if rs_cost.eof or rs_cost.bof then
								new_date = "2015-01-01"
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
								end_month = "����"
								end_view = ""
								batch_yn = ""
								bonbu_yn = ""
								ceo_yn = ""
								batch_view = ""
								ceo_view = ""
								reg_name = ""
								reg_id = ""
								reg_date = ""
							  else														
								new_date = dateadd("m",1,datevalue(mid(rs_cost("end_month"),1,4) + "-" + mid(rs_cost("end_month"),5,2) + "-01"))
								new_month = mid(cstr(new_date),1,4) + mid(cstr(new_date),6,2)
								now_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
									
								if  rs_cost("end_yn") = "Y" then
									end_view = "����"
								  elseif rs_cost("end_yn") = "C" then
									new_month = rs_cost("end_month")
									end_view = "���"
								  else
									end_view = "����"
								end if
								end_yn = rs_cost("end_yn")
								end_month = rs_cost("end_month")							
								reg_name = rs_cost("reg_name")
								reg_id = rs_cost("reg_id")
								reg_date = rs_cost("reg_date")
								if rs_cost("batch_yn") = "Y" then
									batch_view = "�ڷ����"
								  else
								  	batch_view = "�̻���"
								end if
								if rs_cost("bonbu_yn") = "Y" then
									bonbu_view = "���οϷ�"
								end if
								if rs_cost("ceo_yn") = "Y" then
									ceo_view = "���οϷ�"
								end if
								if rs_cost("batch_yn") = "Y" and rs_cost("bonbu_yn") = "N" then
									bonbu_view = "������"
								  	ceo_view = ""
								end if								
								if rs_cost("bonbu_yn") = "Y" and rs_cost("ceo_yn") = "N" then
								  	ceo_view = "������"
								end if								
								if rs_cost("batch_yn") = "N" and rs_cost("bonbu_yn") = "N" and rs_cost("ceo_yn") = "N" then
									bonbu_view = ""
								  	ceo_view = ""
								end if								
								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
							end if					
							if jik_yn = "Y" then
								if ceo_yn = "N" then
									cancel_yn = "Y"
								end if
							  else
							  	if bonbu_yn = "N" then
									cancel_yn = "Y"
								end if
							end if									
						%>
							<tr bgcolor="#99FFFF">
								<td class="first">�κа������</td>
							  	<td><%=end_month%></td>
								<td><%=end_view%>&nbsp;</td>
								<td><%=reg_name%>(<%=reg_id%>)</td>
								<td><%=reg_date%>&nbsp;</td>
							  	<td>
							<% if cancel_yn = "Y" then	%>
                                <a href="company_as_sum_cancel.asp?end_month=<%=end_month%>" class="btnType03">�������</a>
							<%   else	%>
								��ҺҰ�
                            <% end if	%>
                                </td>
								<td><input name="new_month" type="text" id="new_month" style="width:60px; text-align:center" value="<%=new_month%>" readonly="true"></td>
								<td>
							<% if now_month > new_month then	%>
                                <a href="company_as_sum_pro.asp?end_month=<%=new_month%>&end_yn=<%=end_yn%>" class="btnType03">����</a>
							<%   else	%>
								�����Ұ�
                            <% end if	%>
                               	</td> 
								<td colspan="3">&nbsp;</td>
						  </tr>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

