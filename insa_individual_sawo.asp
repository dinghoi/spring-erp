<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_individual_sawo.asp"

curr_date = datevalue(mid(cstr(now()),1,10))

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect


    in_pay_sum = 0
	give_pay_sum = 0

    sql="select * from emp_sawo_mem WHERE sawo_empno = '"+in_empno+"'"
	Rs_sum.Open Sql, Dbconn, 1

	do until rs_sum.eof
	   in_pay_sum = in_pay_sum + rs_sum("sawo_in_pay")
	   give_pay_sum = give_pay_sum + rs_sum("sawo_give_pay")

	   rs_sum.movenext()
	loop
    rs_sum.close()

sql = "select * from emp_sawo_mem WHERE sawo_empno = '"+in_empno+"'"
Rs.Open Sql, Dbconn, 1

title_line = " ����ȸ ���� ��Ȳ "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
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
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.condi.value == "") {
					alert ("�Ҽ��� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psawo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_sawo.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
							<col width="4%" >
                            <col width="9%" >
                            <col width="9%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
                            <col width="5%" >
							<col width="6%" >
							<col width="5%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">�Ҽ�</th>
								<th scope="col">������</th>
								<th scope="col">���Ա���</th>
								<th scope="col">Ż����</th>
                                <th scope="col">Ż�𱸺�</th>
                                <th scope="col">�޿�����</th>
                                <th scope="col">����Ƚ��</th>
                                <th scope="col">���Աݾ�</th>
                                <th scope="col">����Ƚ��</th>
                                <th scope="col">���ޱݾ�</th>
								<th colspan="3" scope="col">��&nbsp;&nbsp;��&nbsp;&nbsp;ȸ</th>
							</tr>
						</thead>
					<tbody>
						<%

						do until rs.eof

		                  sawo_empno = rs("sawo_empno")
		                  sawo_emp_name = rs("sawo_emp_name")

                         if sawo_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&sawo_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
		                      emp_position = Rs_emp("emp_position")
		                   end if
	                       Rs_emp.Close()
	                	 end if
						%>
							<tr>
								<td class="first"><%=rs("sawo_empno")%>&nbsp;</td>
                                <td><%=rs("sawo_emp_name")%>&nbsp;</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rs("sawo_company")%>&nbsp;</td>
                                <td><%=rs("sawo_org_name")%>&nbsp;</td>
                                <td><%=rs("sawo_date")%>&nbsp;</td>
                                <td><%=rs("sawo_id")%>&nbsp;</td>
                                <td><%=rs("sawo_out_date")%>&nbsp;</td>
                                <td><%=rs("sawo_out")%>&nbsp;</td>
                                <% If rs("sawo_target") = "Y" then sawo_target = "����" end if %>
                                <% If rs("sawo_target") = "N" then sawo_target = "����" end if %>
								<td><%=sawo_target%>&nbsp;</td>
                                <td style="text-align:right">
                                <a href="#" onClick="pop_Window('insa_sawo_in_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','sawo_inview','scrollbars=yes,width=800,height=400')"><%=rs("sawo_in_count")%></a>
								</td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_in_pay")),0)%>&nbsp;</td>
                                <td style="text-align:right">
                                <a href="#" onClick="pop_Window('insa_sawo_give_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','sawo_inview','scrollbars=yes,width=1000,height=400')"><%=rs("sawo_give_count")%></a>
                                </td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_give_pay")),0)%>&nbsp;</td>
                                <td colspan="3">
                                <a href="#" onClick="pop_Window('insa_sawo_ask.asp?ask_empno=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&u_type=<%=""%>','insa_sawo_ask_pop','scrollbars=yes,width=750,height=350')">�����ݽ�û</a>&nbsp;</td>
 							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>

                        	<tr>
                              <th colspan="2">�Ѱ�</th>
                              <th colspan="2">&nbsp;</th>
                              <th>�� ���Ծ� :</th>
                              <th class="right"><%=formatnumber(clng(in_pay_sum),0)%></th>
                              <th colspan="2">&nbsp;</th>
                              <th colspan="2">�� ���Ծ� :</th>
                              <th colspan="2" class="right"><%=formatnumber(clng(give_pay_sum),0)%></th>
                              <th>&nbsp;</th>
                              <th>�� �� :</th>
                              <th colspan="2" class="right"><%=formatnumber(clng(in_pay_sum-give_pay_sum),0)%></th>
                              <th colspan="2">&nbsp;</th>
							</tr>

						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <div id="paging">
                        <a href="insa_individual_sawo.asp?page=<%=first_page%>&view_sort=<%=view_sort%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_individual_sawo.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_individual_sawo.asp?page=<%=i%>&view_sort=<%=view_sort%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_individual_sawo.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>">[����]</a> <a href="insa_individual_sawo.asp?page=<%=total_page%>&view_sort=<%=view_sort%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
                    <%' if user_id = "900002"  then
					 if user_id = "102592"  then
					%>
				    <td width="15%">
					<div class="btnCenter">
					<a href="#" onClick="pop_Window('insa_sawo_in_list.asp?sawo_empno=<%=sawo_empno%>&emp_name=<%=sawo_emp_name%>','insa_sawo_in_pop','scrollbars=yes,width=900,height=600')" class="btnType04">����ȸ ȸ�񳻿�</a>
					</div>
                    </td>
				    <td width="15%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_sawo_give_list.asp?sawo_empno=<%=sawo_empno%>&emp_name=<%=sawo_emp_name%>','insa_sawo_give_pop','scrollbars=yes,width=1200,height=600')" class="btnType04">������ ���޳���</a>
					</div>
                    </td>
			      </tr>
                  <% end if %>
				  </table>
			</form>
		</div>
	</div>
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

