<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim win_sw

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	slip_month=Request("slip_month")
  else
	slip_month=Request.form("slip_month")
End if

if slip_month = "" then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
end If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

pgsize = 10 ' ȭ�� �� ������

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select * from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no ='"&user_id&"') and (account_end ='Y') "
order_sql = " ORDER BY slip_date ASC"

sql = "select count(*) from card_slip where (cost_vat > 0) and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no ='"&user_id&"') and (account_end ='Y') "
Set rs_vat = Dbconn.Execute (sql)

vat_record = cint(rs_vat(0)) 'Result.RecordCount

sql = "select count(*) from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_no ='"&user_id&"') and (account_end ='Y') "
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (slip_date >= '" + from_date  + "' and slip_date <= '" + to_date  + "') and (emp_no ='"&user_id&"') and (account_end ='Y') "
Set rs_sum = Dbconn.Execute (sql)
if isnull(rs_sum("price")) then
	sum_price = 0
	sum_cost = 0
	sum_cost_vat = 0
  else
	sum_price = cdbl(rs_sum("price"))
	sum_cost = cdbl(rs_sum("cost"))
	sum_cost_vat = cdbl(rs_sum("cost_vat"))
end if

sql = base_sql + order_sql + " limit "& stpage & "," &pgsize
'response.write(sql)
Rs.Open Sql, Dbconn, 1

title_line = "���κ� ī�� ��� ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {  $( "#datepicker" ).datepicker();
							$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});
			$(function() {  $( "#datepicker1" ).datepicker();
							$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.slip_month.value == "") {
					alert ("������� �Է��ϼ���");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="person_card_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�����&nbsp;</strong>(��201401) :
                                	<input name="slip_month" type="text" value="<%=slip_month%>" style="width:70px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="11%" >
							<col width="13%" >
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
							<col width="10%" >
							<col width="10%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�����</th>
								<th scope="col">ī������</th>
								<th scope="col">ī���ȣ</th>
								<th scope="col">���μ�/�����</th>
								<th scope="col">�ŷ�ó</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">��������</th>
								<th scope="col">�׸�</th>
								<th scope="col">Ȯ��</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<th colspan="2" class="first">�Ѱ�</th>
								<th><%=tottal_record%>&nbsp;��</th>
							  	<th colspan="2"><%=err_msg%>&nbsp;�հ� :&nbsp;<%=formatnumber(sum_price,0)%></th>
							  	<th colspan="3">���ް��� :&nbsp;<%=formatnumber(sum_cost,0)%></th>
								<th colspan="3">�ΰ��� :&nbsp;<%=formatnumber(sum_cost_vat,0)%>&nbsp;(<%=vat_record%>��)</th>
							</tr>
						<%
						i = 0
						j = 0
						person_end = ""
						end_sw = ""
						do until rs.eof

						    person_end = rs("person_end")
						    end_sw = rs("end_sw")
							i = i + 1
							if rs("cost_vat") <> 0 then
								j = j + 1
							end if
							if rs("person_end") = "" or isnull(rs("person_end")) then
								person_end = "N"
							  else
							  	person_end = rs("person_end")
							end if
						%>
							<tr>
								<td class="first"><%=rs("slip_date")%>
                                <input type="hidden" name="approve_no" value="<%=rs("approve_no")%>"></td>
								<td><%=rs("card_type")%></td>
								<td><%=rs("card_no")%></td>
								<td><%=rs("org_name")%>&nbsp;/&nbsp;<%=rs("emp_name")%></td>
								<td><%=rs("customer")%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("account")%>&nbsp;</td>
								<td><%=rs("account_item")%>&nbsp;</td>
								<td>                                
                                <% if person_end = "N" and end_sw = "N" then	%>
                                    <a href="#" onClick="pop_Window('card_slip_mod.asp?approve_no=<%=rs("approve_no")%>&cancel_yn=<%=rs("cancel_yn")%>&person_yn=<%="Y"%>','ī����ǥ����','scrollbars=yes,width=800,height=300')">����</a>
                                <% else %>
                                    <% if person_end = "Y" then %>
                                        Ȯ��
                                    <% else %>
                                        &nbsp;
                                    <% end if %>
                                    <% if end_sw = "Y" then %>
                                        ����
                                    <% end if %>
                                <% end if %>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						if price_sum <> ( cost_sum + cost_vat_sum ) then
							err_msg = "�ݾ�Ȯ�� ���"
						  else
						  	err_msg = " "
						end if
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
                        <a href="person_card_slip_excel.asp?emp_no=<%=user_id%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="person_card_mg.asp?page=<%=first_page%>&slip_month=<%=slip_month%>&ck_sw=<%="y"%>">[ó��]</a>
                        <% if intstart > 1 then %>
                            <a href="person_card_mg.asp?page=<%=intstart -1%>&slip_month=<%=slip_month%>&ck_sw=<%="y"%>">[����]</a>
                        <% end if %>
                        <% for i = intstart to intend %>
                            <% if i = int(page) then %>
                                <b>[<%=i%>]</b>
                            <% else %>
                                <a href="person_card_mg.asp?page=<%=i%>&slip_month=<%=slip_month%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% end if %>
                        <% next %>
                        <% if 	intend < total_page then %>
                            <a href="person_card_mg.asp?page=<%=intend+1%>&slip_month=<%=slip_month%>&ck_sw=<%="y"%>">[����]</a>
                            <a href="person_card_mg.asp?page=<%=total_page%>&slip_month=<%=slip_month%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                            [����]&nbsp;[������]
                        <% end if %>
                    </div>
                    </td>
				    <td width="25%">
					<div class="btnCenter">
                    <% if (person_end = "N") and (i <> 0) and (end_sw = "N") then	%>
                        <a href="person_card_end.asp?slip_month=<%=slip_month%>&emp_no=<%=user_id%>" class="btnType04">����ó��</a>
                    <% end if %>
                    <% if (person_end = "Y") and (i <> 0) and (end_sw = "N") then	%>
                        <a href="person_card_end_cancel.asp?slip_month=<%=slip_month%>&emp_no=<%=user_id%>" class="btnType04">�������ó��</a>
                    <% end if %>
					</div>
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="user_id">
				<input type="hidden" name="pass">
			</form>
		</div>
	</div>
	</body>
</html>
