<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
code_ary = request("code_ary")
goods_type = Request("goods_type") 
stock_code = Request("stock_code") 

ck_sw=Request("ck_sw")
 
'If ck_sw = "y" Then
If goods_type <> "" Then
	code_ary = request("code_ary")
    goods_type = Request("goods_type") 
    stock_code = Request("stock_code") 
  else
	code_ary = Request.form("code_ary")
    goods_type = Request.form("goods_type") 
    stock_code = Request.form("stock_code") 
End if

acpt_standard = Request.form("acpt_standard")
acpt_gubun = Request.form("acpt_gubun")
view_c = Request.form("view_c")

if view_c = "" then
	    acpt_standard = ""
	    acpt_gubun = ""
	    view_c = "acpt"
end if

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

Sql = "SELECT * FROM met_stock_code where stock_code = '"&stock_code&"'"
Set Rs_stock = DbConn.Execute(SQL)
if not Rs_stock.eof then
    	stock_name = Rs_stock("stock_name")
		stock_level = Rs_stock("stock_level")
   else
		stock_name = ""
		stock_level = ""
end if
Rs_stock.close()


'if goods_type = "" then
'	sql = "select * from met_stock_gmaster where stock_code = '" + stock_code + "' and stock_goods_type = '" + goods_type + "'"
'  else
'	sql = "select * from met_stock_gmaster where stock_code = '" + stock_code + "' and stock_goods_type like '%" + goods_type + "%' ORDER BY stock_goods_name,stock_goods_code ASC"
'end if

if view_c = "acpt" then
	sql = "select * from met_stock_gmaster where stock_code = '"&stock_code&"' and stock_goods_type = '"&goods_type&"' and stock_goods_name like '%"&acpt_gubun&"%' ORDER BY stock_goods_gubun,stock_goods_name,stock_goods_code ASC"
  else
	sql = "select * from met_stock_gmaster where stock_code = '"&stock_code&"' and stock_goods_type = '"&goods_type&"' and stock_goods_gubun like '%"&acpt_standard&"%' ORDER BY stock_goods_gubun,stock_goods_name,stock_goods_code ASC"
end if
Rs.Open Sql, Dbconn, 1

title_line = " â�� ��� ǰ�� �˻�"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>ǰ�� �˻�</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function frmcheck1 () {
//				if (chkfrm1()) {
				document.frm1.submit ();
//				}
			}			
			
			function chkfrm() {
				if(document.frm.goods_type.value == "" || document.frm.goods_type.value == " ") {
					alert('�뵵������ �Է��ϼ���');
					frm.goods_type.focus();
					return false;}
				{
					return true;
				}
			}
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('p_gubun').style.display = 'none';
					document.getElementById('p_standard').style.display = '';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('p_gubun').style.display = '';
					document.getElementById('p_standard').style.display = 'none';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_stock_goods_select.asp?code_ary=<%=code_ary%>&ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>�뵵����</strong>
								<label>
                                <input name="goods_type" type="text" value="<%=goods_type%>" readonly="true" style="width:90px" id="goods_type">
								</label>
                                <label>
                              	<input type="radio" name="view_c" value="acpt" <% if view_c = "acpt" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                ǰ���
                                <input type="radio" name="view_c" value="work" <% if view_c = "work" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                ǰ�񱸺�
								</label>
                                <label id="p_standard">
								<strong>ǰ���</strong>
                                    <input name="acpt_gubun" type="text" id="acpt_gubun" value="<%=acpt_gubun%>" style="width:120px">
								</label>
                              <%
								Sql="select * from met_etc_code where etc_type = '04' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>                                       
								<label id="p_gubun">
								<strong>ǰ�񱸺�</strong>
                                <select name="acpt_standard" id="acpt_standard" type="text" style="width:120px">
                                    <option value="">����</option>
                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("etc_name")%>' <%If acpt_standard = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
							  %>
            					</select>
								</label>
                                <label>
								<strong>â��: </strong>
                                	<input name="stock_name" type="text" value="<%=stock_name%>" readonly="true" style="width:100px" id="stock_name">
                                    <input name="stock_level" type="text" value="<%=stock_level%>" readonly="true" style="width:60px" id="stock_level">
                                    <input type="hidden" name="stock_code" value="<%=stock_code%>">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<form action="met_stock_goods_select_ok.asp" method="post" name="frm1">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
                            <col width="10%" >
                            <col width="12%" >
							<col width="*" >
							<col width="16%" >
							<col width="16%" >
                            <col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">����</th>
                                <th scope="col">�뵵����</th>
                                <th scope="col">ǰ���ڵ�</th>
                                <th scope="col">ǰ���</th>
								<th scope="col">ǰ�񱸺�</th>
								<th scope="col">�԰�</th>
                                <th scope="col">������</th>
							</tr>
						</thead>
						<tbody>
						<%
							i = 0
							do until rs.eof or rs.bof
								i = i + 1
						%>
							<tr>
								<td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="<%=rs("stock_goods_code")%>"></td>
								<td><%=rs("stock_goods_grade")%>&nbsp;</td>
                                <td><%=rs("stock_goods_type")%>&nbsp;</td>
                                <td><%=rs("stock_goods_code")%>&nbsp;</td>
                                <td><%=rs("stock_goods_name")%>&nbsp;</td>
                                <td><%=rs("stock_goods_gubun")%>&nbsp;</td>
								<td><%=rs("stock_goods_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("stock_JJ_qty"),0)%>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						<%
						  if i = 0 then
						%>
							<tr>
								<td class="first" colspan="8">������ �����ϴ�</td>
							</tr>
                        <%
						end if
						%>
							<tr>
								<td class="first; left" colspan="8"><span class="btnType04"><input type="button" value="����" onclick="javascript:frmcheck1();"></span></td>
							</tr>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="code_ary" value="<%=code_ary%>">
                <input type="hidden" name="stock_code1" value="<%=stock_code%>">
                <input type="hidden" name="goods_type1" value="<%=goods_type%>">
				</form>
		</div>        				
	</body>
</html>

