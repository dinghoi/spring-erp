<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
approve_no = request("approve_no")
'saupbu = request("saupbu")

'if saupbu = "-" or saupbu = "" then
'	saupbu = "����οܳ�����"
'end if

'sql = "select max(end_month) as max_month from cost_end where saupbu = '"&saupbu&"' and end_yn ='Y'"
'set rs_end=dbconn.execute(sql)
'if	isnull(rs_end("max_month")) then
'	end_date = "2014-08-31"
'  else
'	new_date = dateadd("m",1,datevalue(mid(rs_end("max_month"),1,4) + "-" + mid(rs_end("max_month"),5,2) + "-01"))
'	end_date = dateadd("d",-1,new_date)
'end if
'rs_end.close()
'if saupbu = "����οܳ�����" then
'	saupbu = ""
'end if

Sql="select * from tax_bill where approve_no = '"&approve_no&"'"
Set rs=DbConn.Execute(Sql)

Sql="select * from trade where trade_no = '"&rs("trade_no")&"'"
Set rs_trade=DbConn.Execute(Sql)
if rs_trade.eof or rs_trade.bof then
	customer = rs("trade_name")
  else
	customer = rs_trade("trade_name")
end if
emp_no = user_id
emp_name = user_name
emp_grade = user_grade
account = ""
end_yn = "N"
curr_date = mid(cstr(now()),1,10)
mg_saupbu = "����"

owner_company = rs("owner_company")
Select Case owner_company
	Case "���̿��������" : owner_company = "���̿�"
	Case "�ڸ��Ƶ𿣾�" : owner_company = "���̽ý���"
	Case Else
		owner_company = rs("owner_company")
End Select

'������� ����
SQL = "SELECT org_name "
SQL = SQL & "FROM emp_master AS emtt "
SQL = SQL & "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
SQL = SQL & "WHERE emtt.emp_no = '"&emp_no&"' "

Set rsOrg = DBConn.Execute(SQL)

org_name = rsOrg("org_name")

rsOrg.Close() : Set rsOrg = Nothing

title_line = "E���� ���� ���ݰ�꼭 ���κ�� �Է�"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
//				if(document.frm.bill_date.value <= document.frm.end_date.value) {
//					alert('�������ڰ� ������ �Ǿ� �ִ� �����Դϴ�');
//					frm.slip_date.focus();
//					return false;}
				if(document.frm.mg_saupbu.value =="����") {
					alert('��翵������θ� �����ϼ���');
					frm.mg_saupbu.focus();
					return false;}
				if(document.frm.company.value =="") {
					alert('���縦 �����ϼ���');
					frm.company.focus();
					return false;}
				if(document.frm.slip_gubun.value =="") {
					alert('��������� �����ϼ���');
					frm.slip_gubun.focus();
					return false;}

				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_esero_in_detail_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">��������</th>
				        <td class="left"><%=rs("bill_date")%></td>
				        <th>���޹޴�ȸ��</th>
				        <td class="left"><%'=rs("owner_company")%><%=owner_company%></td>
			          </tr>
				      <tr>
				        <th class="first">�������</th>
				        <td class="left">
                        <% if cost_grade = "0" or saupbu = "�濵������" or team = "SM1��" or team = "SM2��" or team = "Repair��" or user_name = "��ȣ��" Or user_id = "101756" then	%>
                            <input name="org_name" type="text" readonly="true" value="<%=org_name%>" style="width:150px">
                            <a href="#" onClick="pop_Window('/org_search.asp?gubun=<%="��꼭"%>&org_company=<%'=rs("owner_company")%><%=owner_company%>','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
                        <% else	%>
                            <%=org_name%>
                            <input name="org_name" type="hidden" value="<%=org_name%>">
                        <% end if %>
				        <input name="emp_company" type="hidden" value="<%=emp_company%>">
				        <input name="bonbu" type="hidden" value="<%=bonbu%>">
				        <input name="saupbu" type="hidden" value="<%=saupbu%>">
				        <input name="team" type="hidden" value="<%=team%>">
				        <input name="reside_place" type="hidden" value="<%=reside_place%>">
                        <input name="reside_company" type="hidden" value="<%=reside_company%>">
                        </td>
				        <th>��翵�������</th>
				        <td class="left">
							<%
                            cost_year = mid(rs("bill_date"),1,4)
                            sql_org = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
                            rs_org.Open sql_org, Dbconn, 1
                            'Response.write sql_org
                            %>
                            <select name="mg_saupbu" id="mg_saupbu" style="width:150px">
                                <option value='����' <%If mg_saupbu = "����" then %>selected<% end if %>>����</option>
                                <option value='' <%If mg_saupbu = "" then %>selected<% end if %>>��翵���ξ���</option>
                                <%
                                do until rs_org.eof
                                    %>
                                    <option value='<%=rs_org("saupbu")%>' <%If rs_org("saupbu") = mg_saupbu  then %>selected<% end if %>><%=rs_org("saupbu")%></option>
                                    <%
                                    rs_org.movenext()
                                    loop
                                rs_org.Close()
                                %>
                            </select>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">������</th>
				        <td class="left"><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%>&nbsp;<%=rs("trade_name")%></td>
				        <th>�����</th>
				        <td class="left"><input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=emp_name%>" readonly="true">
                          <input name="emp_grade" type="text" id="emp_grade" style="width:60px" value="<%=emp_grade%>" readonly="true">
                        <a href="#" onClick="pop_Window('/insa/emp_search.asp?gubun=<%="1"%>','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">�����ȸ</a></td>
			          </tr>
				      <tr>
				        <th class="first">���೻��</th>
				        <td class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=rs("tax_bill_memo")%>"></td>
				        <th>�ݾ�</th>
				        <td class="left"><strong>���ް��� :</strong>&nbsp;<%=formatnumber(rs("cost"),0)%>&nbsp;&nbsp;&nbsp;<strong>�ΰ��� :</strong>&nbsp;<%=formatnumber(rs("cost_vat"),0)%></td>
			          </tr>
				      <tr>
				        <th class="first">����</th>
				        <td class="left">
                        <input name="company" type="text" value="<%=company%>" readonly="true" style="width:150px">
			            <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
                        </td>
				        <th>�������</th>
				        <td class="left">
						<input type="text" name="slip_gubun" ID="slip_gubun" readonly="true" style="width:100px">
						<input name="account_view" type="text" readonly="true" style="width:150px">
                        <a href="#" onClick="pop_Window('tax_bill_account_search.asp','tax_bill_account_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
						<input name="account" type="hidden" id="account">
						<input name="account_item" type="hidden" id="account_item">
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <% if end_yn = "N" then	%>
                        <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <% end if %>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
				<input type="hidden" name="bill_date" value="<%=rs("bill_date")%>" ID="Hidden1">
				<input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
				<input type="hidden" name="approve_no" value="<%=approve_no%>" ID="Hidden1">
			</form>
		</div>
	</body>
</html>

