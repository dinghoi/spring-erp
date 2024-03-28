<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	slip_month=Request("slip_month")
	emp_yn=Request("emp_yn")
	emp_name=Request("emp_name")
	sort_condi=Request("sort_condi")
  else
	slip_month=Request.form("slip_month")
	emp_yn=Request.form("emp_yn")
	emp_name=Request.form("emp_name")
	sort_condi=Request.form("sort_condi")
End if

if slip_month = "" then
	be_date = dateadd("m",-1,now())
	slip_month = mid(cstr(be_date),1,4) + mid(cstr(be_date),6,2)
	emp_yn = "N"
	emp_name = ""
	sort_condi = "emp"
End If

If emp_yn = "N" Then
	emp_name = ""
End If

from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))
be_from_date = dateadd("m",-1,from_date)
be_to_date = mid(be_from_date,1,4) + "-" + mid(be_from_date,6,2) + "-31"

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

' ���� ��ȸ
if emp_yn = "Y" then
	condi_sql = " and emp_name like '%"&emp_name&"%'"
  else
  	condi_sql = ""
end if
' ��ȸ����
if sort_condi = "emp" then
	order_sql = " order by card_slip.emp_name asc"
  else
  	order_sql = " order by price desc"
end if

' ���ڵ� �Ǽ�
tottal_record = 0
sql = "select emp_no from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"&condi_sql&" group by emp_no"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	tottal_record = tottal_record + 1
	rs.movenext()
loop
rs.close()		
'tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

' ��� �ݾ� SUM ó��
sql = "select count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"&condi_sql
Set rs = Dbconn.Execute (sql)
sum_cnt = cdbl(rs("slip_cnt")) 
if rs("price") = "" or isnull(rs("price")) then
	sum_cost = 0
	sum_cost_vat = 0
	sum_price = 0
  else
	sum_cost = cdbl(rs("cost")) 
	sum_cost_vat = cdbl(rs("cost_vat")) 
	sum_price = cdbl(rs("price")) 
end if
rs.close()

' ���� �ݾ� ��ü SUM ó��
sql = "select count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (slip_date >='"&be_from_date&"' and slip_date <='"&be_to_date&"')"&condi_sql
Set rs_etc = Dbconn.Execute (sql)
be_sum_cnt = cdbl(rs_etc("slip_cnt")) 
if rs_etc("price") = "" or isnull(rs_etc("price")) then
	be_sum_cost = 0
	be_sum_cost_vat = 0
	be_sum_price = 0
  else
	be_sum_cost = cdbl(rs_etc("cost")) 
	be_sum_cost_vat = cdbl(rs_etc("cost_vat")) 
	be_sum_price = cdbl(rs_etc("price")) 
end if
rs_etc.close()

sql = "select card_slip.emp_no,card_slip.emp_name,memb.user_grade,memb.org_name,count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat "
sql = sql + " from card_slip inner join memb on card_slip.emp_no=memb.user_id where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"&condi_sql&" group by card_slip.emp_no " + order_sql + " limit "& stpage & "," &pgsize 
'response.write(sql)
Rs.Open Sql, Dbconn, 1
'Response.write sql

title_line = "ī�� ��ǥ ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>����ȸ��ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
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
			function condi_view() {

				if (eval("document.frm.emp_yn[0].checked")) {
					document.getElementById('emp_name_view').style.display = 'none';
				}	
				if (eval("document.frm.emp_yn[1].checked")) {
					document.getElementById('emp_name_view').style.display = '';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/card_slip_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="card_person_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�����&nbsp;</strong>(��201401) : 
                                	<input name="slip_month" type="text" value="<%=slip_month%>" style="width:60px">
								</label>
                                <label>
								<strong>�˻�����</strong>
                                  <input type="radio" name="emp_yn" value="N" <% if emp_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio1" onClick="condi_view()">��ü </label>
                                  <input type="radio" name="emp_yn" value="Y" <% if emp_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio2" onClick="condi_view()">������
                                </label>
								&nbsp;&nbsp;
                                <label>
                                	<input name="emp_name" type="text" value="<%=emp_name%>" style="width:80px; display:none" id="emp_name_view">
								</label>
                                <label>
								<strong>��ȸ����</strong>
                                  <input type="radio" name="sort_condi" value="emp" <% if sort_condi = "emp" then %>checked<% end if %> style="width:30px" id="Radio1">������
                                  <input type="radio" name="sort_condi" value="price" <% if sort_condi = "price" then %>checked<% end if %> style="width:30px" id="Radio2">�ݾ׼�
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="5%" >
							<col width="8%" >
							<col width="7%" >
							<col width="8%" >
							<col width="5%" >
							<col width="8%" >
							<col width="7%" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">������</th>
								<th rowspan="2" scope="col">������</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� ��</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">���</th>
								<th rowspan="2" scope="col">������</th>
								<th rowspan="2" scope="col">������</th>
								<th rowspan="2" scope="col">�������<p>������ȸ</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">�Ǽ�</th>
							  <th scope="col">���ް���</th>
							  <th scope="col">�ΰ���</th>
							  <th scope="col">�հ�</th>
							  <th scope="col">�Ǽ�</th>
							  <th scope="col">���ް���</th>
							  <th scope="col">�ΰ���</th>
							  <th scope="col">�հ�</th>
						  </tr>
						</thead>
						<tbody>
							<tr>
								<th class="first">�Ѱ�</th>
								<th><%=formatnumber(tottal_record,0)%>&nbsp;��</th>
							  	<th class="right"><%=formatnumber(be_sum_cnt,0)%></th>
							  	<th class="right"><%=formatnumber(be_sum_cost,0)%></th>
								<th class="right"><%=formatnumber(be_sum_cost_vat,0)%></th>
							  	<th class="right"><%=formatnumber(be_sum_price,0)%></th>
							  	<th class="right"><%=formatnumber(sum_cnt,0)%></th>
							  	<th class="right"><%=formatnumber(sum_cost,0)%></th>
								<th class="right"><%=formatnumber(sum_cost_vat,0)%></th>
							  	<th class="right"><%=formatnumber(sum_price,0)%></th>
							  	<th class="right"><%=formatnumber(sum_price-be_sum_price,0)%></th>
								<th class="right"><%=formatnumber((sum_price-be_sum_price)/be_sum_price*100,2)%>%</th>
								<th>&nbsp;</th>
							</tr>
						<%
						do until rs.eof
' ���� �ݾ� ��ü SUM ó��
							sql = "select count(*) as slip_cnt,sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from card_slip where (slip_date >='"&be_from_date&"' and slip_date <='"&be_to_date&"') and emp_no = '"&rs("emp_no")&"'"
							Set rs_etc = Dbconn.Execute (sql)
							be_cnt = cdbl(rs_etc("slip_cnt")) 
							if rs_etc("price") = "" or isnull(rs_etc("price")) then
								be_cost = 0
								be_cost_vat = 0
								be_price = 0
							  else
								be_cost = cdbl(rs_etc("cost")) 
								be_cost_vat = cdbl(rs_etc("cost_vat")) 
								be_price = cdbl(rs_etc("price")) 
							end if
							if be_price = 0 then
								incr_per = 100
							  else
								incr_per = (cdbl(rs("price"))-be_price)/be_price*100
							end if
						%>
							<tr>
								<td class="first"><%=rs("emp_name")%>&nbsp;<%=rs("user_grade")%></td>
								<td class="left"><%=rs("org_name")%></td>
							  	<td class="right"><%=formatnumber(be_cnt,0)%></td>
							  	<td class="right"><%=formatnumber(be_cost,0)%></td>
							  	<td class="right"><%=formatnumber(be_cost_vat,0)%></td>
							  	<td class="right"><%=formatnumber(be_price,0)%></td>
							  	<td class="right"><%=formatnumber(rs("slip_cnt"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(cdbl(rs("price"))-be_price,0)%></td>
							  	<td class="right"><%=formatnumber(incr_per,2)%>%</td>
								<td>
                               	<input type="hidden" name="emp_no" value="rs("emp_no")"%>                              
								<a href="#" onClick="pop_Window('person_card_slip_view.asp?slip_month=<%=slip_month%>&emp_no=<%=rs("emp_no")%>','ī����ǥ����','scrollbars=yes,width=900,height=500')">��ȸ</a>
                                </td>
							</tr>
						<%
							rs_etc.close()
							rs.movenext()
						loop
						rs.close()
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
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="card_person_mg.asp?page=<%=first_page%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="card_person_mg.asp?page=<%=intstart -1%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="card_person_mg.asp?page=<%=i%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="card_person_mg.asp?page=<%=intend+1%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[����]</a> 
                        <a href="card_person_mg.asp?page=<%=total_page%>&slip_month=<%=slip_month%>&emp_yn=<%=emp_yn%>&emp_name=<%=emp_name%>&sort_condi=<%=sort_condi%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="25%">
					<div class="btnCenter">
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

