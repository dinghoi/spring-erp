<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
sign_date=Request("sign_date")
sign_seq=Request("sign_seq")
sign_month=Request("sign_month")
msg_seq=int(Request("msg_seq"))
sign_pro=Request("sign_pro")
sign_yn=Request("sign_yn")
sign_head=Request("sign_head")
paper_no = sign_date + "-" + sign_seq

from_date = cstr(mid(sign_month,1,4)) + "-" + cstr(mid(sign_month,5,2)) + "-" + "01"
to_date = cstr(mid(sign_month,1,4)) + "-" + cstr(mid(sign_month,5,2)) + "-" + "31"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_sign = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from sign_process where sign_date = '"&sign_date&"' and sign_seq = '"&sign_seq&"'"
Set rs_sign=DbConn.Execute(Sql)

team_emp_name = " "
if rs_sign("team_sign") = "C" then
	team_emp_name = "�ݷ�"
end if
if rs_sign("team_sign") = "E" then
	sql = "select * from memb where user_id = '"&rs_sign("team_emp_no")&"'"
	Set rs_memb=DbConn.Execute(Sql)
	team_emp_name = rs_memb("user_name")
	rs_memb.close()
end if

saupbu_emp_name = " "
if rs_sign("saupbu_sign") = "C" then
	saupbu_emp_name = "�ݷ�"
end if
if rs_sign("saupbu_sign") = "E" then
	sql = "select * from memb where user_id = '"&rs_sign("saupbu_emp_no")&"'"
	Set rs_memb=DbConn.Execute(Sql)
	saupbu_emp_name = rs_memb("user_name")
	rs_memb.close()
end if

bonbu_emp_name = " "
if rs_sign("bonbu_sign") = "C" then
	bonbu_emp_name = "�ݷ�"
end if
if rs_sign("bonbu_sign") = "E" then
	sql = "select * from memb where user_id = '"&rs_sign("bonbu_emp_no")&"'"
	Set rs_memb=DbConn.Execute(Sql)
	bonbu_emp_name = rs_memb("user_name")
	rs_memb.close()
end if

ceo_emp_name = " "
if rs_sign("ceo_sign") = "C" then
	ceo_emp_name = "�ݷ�"
end if
if rs_sign("ceo_sign") = "E" then
	sql = "select * from memb where user_id = '"&rs_sign("ceo_emp_no")&"'"
	Set rs_memb=DbConn.Execute(Sql)
	ceo_emp_name = rs_memb("user_name")
	rs_memb.close()
end if

sql = "update sign_msg set read_yn='Y' where msg_seq="&msg_seq
dbconn.execute(sql)	  

' ���Ǻ� ��ȸ.........
'posi_sql = " and reg_id = '" + user_id + "'"

'if position = "����" and view_c = "total" then
'	posi_sql = " and team = '"&team&"'"
'end if

'base_sql = "select * from general_cost where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
'order_sql = " ORDER BY slip_date ASC"
sql = "select * from general_cost where paper_no ='"&paper_no&"' ORDER BY slip_date ASC"

'sql = base_sql + posi_sql + order_sql
Rs.Open Sql, Dbconn, 1

sub_title_line = ". �ۼ��� : " + rs_sign("reg_user") + "( " + rs_sign("reg_id") + " )"
paper_title = "������ȣ : " + cstr(rs_sign("sign_date")) + "-" + rs_sign("sign_seq")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
//				self.opener.location.reload();
				window.close () ;
			}
			function frmcheck () {
				document.frm.sign_yn.value = "Y"; 
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function frmcheck1 () {
				document.frm.sign_yn.value = "C"; 
				if (formcheck(document.frm) && chkfrm1()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				{
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function chkfrm1() {
				{
				a=confirm('�ݷ��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //�Ӹ��� ����
                factory.printing.footer = ""; //������ ����
                factory.printing.portrait = false; //��¹��� ����: true - ����, false - ����
                factory.printing.leftMargin = 13; //���� ���� ����
                factory.printing.topMargin = 10; //���� ���� ����
                factory.printing.rightMargin = 13; //�����P ���� ����
                factory.printing.bottomMargin = 15; //�ٴ� ���� ����
        //		factory.printing.SetMarginMeasure(2); //�׵θ� ���� ������ ������ ��ġ�� ����
        //		factory.printing.printer = ""; //������ �� ������ �̸�
        //		factory.printing.paperSize = "A4"; //��������
        //		factory.printing.pageSource = "Manusal feed"; //���� �ǵ� ���
        //		factory.printing.collate = true; //������� ����ϱ�
        //		factory.printing.copies = "1"; //�μ��� �ż�
        //		factory.printing.SetPageRange(true,1,1); //true�� �����ϰ� 1,3�̸� 1���� 3������ ���
        //		factory.printing.Printer(true); //����ϱ�
                factory.printing.Preview(); //�����츦 ���ؼ� ���
                factory.printing.Print(false); //�����츦 ���ؼ� ���
            }
		</script>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="form_wrap">			
			<div id="container">
				<br>
				<h3 class="stit"><%=paper_title%></h3>
				<form action="cost_sign_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col"><h3 class="tit" style="color:#F60"><%=sign_head%></h3></th>
								<th scope="col">����</th>
								<th scope="col">�������</th>
								<th scope="col">������</th>
								<th scope="col">��ǥ�̻�</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="first"><h3 class="stit"><%=sub_title_line%></h3></td>
								<td><%=team_emp_name%>&nbsp;</td>
								<td><%=saupbu_emp_name%>&nbsp;</td>
								<td><%=bonbu_emp_name%>&nbsp;</td>
								<td><%=ceo_emp_name%>&nbsp;</td>
							</tr>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="11%" >
							<col width="8%" >
							<col width="8%" >
							<col width="6%" >
							<col width="7%" >
							<col width="6%" >
							<col width="*" >
							<col width="8%" >
							<col width="5%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�߻�����</th>
								<th scope="col">�Ҽ�</th>
								<th scope="col">�׸�</th>
								<th scope="col">���ڰ���NO</th>
								<th scope="col">��뱸��</th>
								<th scope="col">�ݾ�</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">�߻�����/�ŷ�ó</th>
								<th scope="col">�����</th>
								<th scope="col">����</th>
								<th scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							if rs("pay_yn") = "Y" then
								pay_yn = "����"
							  else
							  	pay_yn = "������"
							end if
							if rs("end_yn") = "Y" then
								end_yn = "����"
								end_view = "N"
							  else
							  	end_yn = "����"
							end if
							belong = rs("team") + " " + rs("belong")
							if rs("team") = "" then
								belong = rs("saupbu")
							end if
							if belong = "" then
								belong = rs("bonbu")
							end if
						%>
							<tr>
								<td class="first"><%=rs("slip_date")%></td>
								<td><%=belong%></td>
								<td><%=rs("account_item")%></td>
								<td><%=rs("paper_no")%></td>
								<td><%=rs("pay_method")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("customer")%></td>
								<td><%=rs("use_man")%><%=rs("user_grade")%>&nbsp;</td>
								<td><%=pay_yn%></td>
								<td><%=rs("slip_memo")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr>
								<th class="first" colspan="5">�� ��</th>
							  	<th class="right"><%=formatnumber(cost_sum,0)%></th>
							  	<th class="right"><%=formatnumber(cost_vat_sum,0)%></th>
							  	<th class="right" colspan="4">&nbsp;</th>
							</tr>
							<tr>
								<td class="first" bgcolor="#CCFFFF">Ư�̻���</td>
						  	  <td class="left" colspan="10"><textarea name="sign_memo" cols="140" rows="3" id="textarea"><%=rs_sign("sign_memo")%></textarea></td>
						  </tr>
						</tbody>
					</table>
				</div>
   				<div class="noprint">
				<br>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
		            <div align=center>
				<% if sign_yn = "N" then	%>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1"></span>
                    <span class="btnType01"><input type="button" value="�ݷ�" onclick="javascript:frmcheck1();" ID="Button1"></span>
				<% end if	%>
               		<span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>            
                    <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
					</div>                  
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="sign_date" value="<%=sign_date%>" ID="Hidden1">
				<input type="hidden" name="sign_seq" value="<%=sign_seq%>" ID="Hidden1">
				<input type="hidden" name="sign_month" value="<%=sign_month%>" ID="Hidden1">
				<input type="hidden" name="msg_seq" value="<%=msg_seq%>" ID="Hidden1">
				<input type="hidden" name="sign_yn" value="<%=sign_yn%>" ID="Hidden1">
				<input type="hidden" name="title_line" value="<%=title_line%>" ID="Hidden1">
				<br>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

