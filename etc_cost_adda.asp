<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
u_type = request("u_type")
slip_date = request("slip_date")
slip_seq = request("slip_seq")

customer = ""
emp_company = ""
org_name = ""
company = ""
account = ""
price = 0
cost = 0
cost_vat = 0
slip_memo = ""
end_yn = "N"
curr_date = mid(cstr(now()),1,10)

title_line = "���� �� ��� ��� ���"
if u_type = "U" then

	Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
	Set rs=DbConn.Execute(Sql)

	customer = rs("customer_no")
	emp_company = rs("emp_company")
	bonbu = rs("bonbu")
	saupbu = rs("saupbu")
	team = rs("team")
	org_name = rs("org_name")
	company = rs("company")
	account = rs("account")
	price = rs("price")
	cost = rs("cost")
	cost_vat = rs("cost_vat")
	slip_memo = rs("slip_memo")
	rs.close()

	title_line = "���� �� ��� ��� ����"
end if
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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=slip_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.slip_date.value <= document.frm.end_date.value) {
					alert('�߻����ڰ� ������ �Ǿ� �ִ� �����Դϴ�');
					frm.slip_date.focus();
					return false;}
				if(document.frm.slip_date.value > document.frm.curr_date.value) {
					alert('�߻����ڰ� �����Ϻ��� Ŭ���� �����ϴ�.');
					frm.slip_date.focus();
					return false;}
				if(document.frm.end_yn.value =="Y") {
					alert('�����Ǿ� ���� �� �� �����ϴ�');
					frm.end_yn.focus();
					return false;}
				if(document.frm.slip_date.value =="") {
					alert('�߻����ڸ� �Է��ϼ���');
					frm.slip_date.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('�ŷ�ó �����ϼ���');
					frm.customer.focus();
					return false;}
				if(document.frm.org_name.value =="") {
					alert('��������� �����ϼ���');
					frm.org_name.focus();
					return false;}
				if(document.frm.company.value =="") {
					alert('���ȸ�縦 �����ϼ���');
					frm.company.focus();
					return false;}
				if(document.frm.price.value ==0) {
					alert('�հ�ݾ��� �Է��ϼ���');
					frm.price.focus();
					return false;}
				if(document.frm.cost_vat.value ==0) {
					alert('�ΰ����� �Է��ϼ���');
					frm.cost_vat.focus();
					return false;}
				if(document.frm.account.value =="") {
					alert('���������� �����ϼ���');
					frm.account.focus();
					return false;}
				if(document.frm.slip_memo.value =="") {
					alert('���೻���� �Է��ϼ���');
					frm.slip_memo.focus();
					return false;}

				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function cost_cal(txtObj){
				price = parseInt(document.frm.price.value.replace(/,/g,""));			
				cost_vat = parseInt(document.frm.cost_vat.value.replace(/,/g,""));			
				cost = price - cost_vat;
				cost = String(cost);
				num_len = cost.length;
				sil_len = num_len;
				cost = String(cost);
				if (cost.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) cost = cost.substr(0,num_len -3) + "," + cost.substr(num_len -3,3);
				if (sil_len > 6) cost = cost.substr(0,num_len -6) + "," + cost.substr(num_len -6,3) + "," + cost.substr(num_len -2,3);

				document.frm.cost.value = cost; 

				if (txtObj.value.length >= 2) {
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}
				if (txtObj.value.length<5) {
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
				}
				var num = txtObj.value;
				if (num == "--" ||  num == "." ) num = "";
				if (num != "" ) {
					temp=new String(num);
					if(temp.length<1) return "";
					
					// ����ó��
					if(temp.substr(0,1)=="-") minus="-";
					else minus="";
					
					// �Ҽ�������ó��
					dpoint=temp.search(/\./);
					
					if(dpoint>0)
					{
					// ù��° ������ .�� �������� �ڸ��� ���������� ���� ����
					dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
					temp=temp.substr(0,dpoint);
					}else dpointVa="";
					
					// �����ܹ̿��� ����
					temp=temp.replace(/\D/g,"");
					zero=temp.search(/[1-9]/);
					
					if(zero==-1) return "";
					else if(zero!=0) temp=temp.substr(zero);
					
					if(temp.length<4) return minus+temp+dpointVa;
					buf="";
					while (true)
					{
					if(temp.length<3) { buf=temp+buf; break; }
				
					buf=","+temp.substr(temp.length-3)+buf;
					temp=temp.substr(0, temp.length-3);
					}
					if(buf.substr(0,1)==",") buf=buf.substr(1);
				
					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;
				}else txtObj.value = "0";					
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="outside_cost_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="15%" >
				      <col width="35%" >
				      <col width="15%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">��������</th>
				        <td class="left">
                        <input name="slip_date" type="text" value="<%=slip_date%>" style="width:80px;text-align:center" id="datepicker">
				          ������ : <%=end_date%>
				        <input name="curr_date" type="hidden" value="<%=curr_date%>">
				        <input name="slip_seq" type="hidden" value="<%=slip_seq%>">
                        </td>
				        <th>�ŷ�ó</th>
				        <td class="left">
                        <select name="customer" id="customer" style="width:150px">
				          <option value="" <% if customer = "" then %>selected<% end if %>>����</option>
				          <%
                            Sql="select * from trade where trade_id = '����' or trade_id = '����' order by trade_name asc"
                            rs_trade.Open Sql, Dbconn, 1
                            do until rs_trade.eof
							%>
				          <option value='<%=rs_trade("trade_no")%>' <%If customer = rs_trade("trade_no") then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
				          <%
                            	rs_trade.movenext()
                            loop
                            rs_trade.close()						
                            %>
			            </select></td>
			          </tr>
				      <tr>
				        <th class="first">�������</th>
				        <td class="left">
                          <a href="#" onClick="pop_Window('org_search.asp','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">������ȸ</a>
                          <input name="org_name" type="text" value="<%=org_name%>" readonly="true" style="width:150px">
                          <input name="emp_company" type="hidden" value="<%=emp_company%>">
                          <input name="bonbu" type="hidden" value="<%=bonbu%>">
                          <input name="saupbu" type="hidden" value="<%=saupbu%>">
                          <input name="team" type="hidden" value="<%=team%>">
                          <input name="reside_place" type="hidden" value="<%=reside_place%>">
                        </td>
				        <th>���ȸ��</th>
				        <td class="left"><select name="company" id="company" style="width:150px">
				          <option value="" <% if company = "" then %>selected<% end if %>>����</option>
				          <option value="����" <% if company = "����" then %>selected<% end if %>>����</option>
				          <%
                            Sql="select * from trade where trade_id = '����' or trade_id = '����' order by trade_name asc"
                            rs_trade.Open Sql, Dbconn, 1
                            do until rs_trade.eof
							%>
				          <option value='<%=rs_trade("trade_name")%>' <%If company = rs_trade("trade_name") then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
				          <%
                            	rs_trade.movenext()
                            loop
                            rs_trade.close()						
                            %>
			            </select></td>
			          </tr>
				      <tr>
				        <th class="first">�հ�ݾ�</th>
				        <td class="left"><input name="price" type="text" id="price" style="width:100px;text-align:right" value="<%=formatnumber(price,0)%>" onKeyUp="cost_cal(this);" ></td>
				        <th><span class="first">�ΰ���</span></th>
				        <td class="left"><input name="cost_vat" type="text" id="cost_vat" style="width:100px;text-align:right" value="<%=formatnumber(cost_vat,0)%>" onKeyUp="cost_cal(this);" ></td>
			          </tr>
				      <tr>
				        <th class="first">���ް���</th>
				        <td class="left"><input name="cost" type="text" id="cost" style="width:100px;text-align:right" value="<%=formatnumber(cost,0)%>" readonly="true" ></td>
				        <th>��������</th>
				        <td class="left">
						<select name="account" id="account" style="width:150px">
				          <option value="" <% if account = "" then %>selected<% end if %>>����</option>
				          <%
                            Sql="select * from etc_code where etc_type = '43' order by etc_name asc"
                            rs_etc.Open Sql, Dbconn, 1
                            do until rs_etc.eof
							%>
				          <option value='<%=rs_etc("etc_name")%>' <%If account = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
				          <%
                            	rs_etc.movenext()
                            loop
                            rs_etc.close()						
                            %>
			            </select>                        
                        </td>
			          </tr>
				      <tr>
				        <th class="first">���೻��</th>
				        <td colspan="3" class="left"><input name="slip_memo" type="text" id="slip_memo" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=slip_memo%>"></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	if end_yn = "N" then	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
				<input type="hidden" name="old_date" value="<%=slip_date%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

