<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim etc_code_last

ck_sw=Request("ck_sw")
u_type = request("u_type")
card_upjong = request("card_upjong")

If ck_sw = "y" Then
	view_c=Request("view_c")
	field_view=Request("field_view")
  else
	view_c=Request.form("view_c")
	field_view=Request.form("field_view")
End if

if view_c = "" then
	view_c = "total"
	field_view = ""
end if

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")					
Set Rs_etc = Server.CreateObject("ADODB.Recordset")					
DbConn.Open dbconnect

if view_c = "total" then
	sql = "select * from card_upjong order by card_upjong asc"
  elseif view_c = "upjong" then
	sql = "select * from card_upjong where card_upjong like '%"&field_view&"%' order by card_upjong asc"
  else
	sql = "select * from card_upjong where account like '%"&field_view&"%' order by card_upjong asc"
end if
Rs.Open Sql, Dbconn, 1

if u_type = "U" then
	sql = "select * from card_upjong where card_upjong = '" + card_upjong + "'"
	Set rs_etc=DbConn.Execute(Sql)
	account = rs_etc("account")
	account_item = rs_etc("account_item")
	account_view = account + "-" + account_item
	tax_yn = rs_etc("tax_yn")
  else
	account = ""
	account_item = ""
	group_name = ""
	tax_yn = "Y"
end if	

title_line = "카드 거래처 업종 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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
			function frmsubmit () {
				document.condi_frm.submit ();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
			
				if(document.frm.card_upjong.value =="") {
					alert('업종명을 입력하세요');
					frm.card_upjong.focus();
					return false;}
				if(document.frm.account_view.value =="") {
					alert('계정과목을 선택하세요');
					frm.account_view.focus();
					return false;}
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.tax_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("과세구분을 선택하시기 바랍니다");
					return false;
				}	

				a=confirm('등록하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			
			}
			function condi_view() {

				if (eval("document.condi_frm.view_c[0].checked")) {
					document.getElementById('field_view').style.display = 'none';
				}	
				if (eval("document.condi_frm.view_c[1].checked")) {
					document.getElementById('field_view').style.display = '';
				}	
				if (eval("document.condi_frm.view_c[2].checked")) {
					document.getElementById('field_view').style.display = '';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view();">
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/card_slip_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="card_cust_upjong_mg.asp" method="post" name="condi_frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>조회조건 : </strong>
                              	<input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">전체
                                <input type="radio" name="view_c" value="upjong" <% if view_c = "upjong" then %>checked<% end if %> style="width:25px" onClick="condi_view()">업종명
                                <input type="radio" name="view_c" value="account" <% if view_c = "account" then %>checked<% end if %> style="width:25px" onClick="condi_view()">계정과목
								</label>
								<label>
                                	<input name="field_view" type="text" value="<%=field_view%>" style="width:70px; display:none" id="field_view">
								</label>
                                <a href="#" onclick="javascript:frmsubmit();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="64%" height="356" valign="top"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="*" >
				          <col width="20%" >
				          <col width="20%" >
				          <col width="20%" >
				          <col width="10%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">업종명</th>
				            <th scope="col">계정과목</th>
				            <th scope="col">항목</th>
				            <th scope="col">과세구분</th>
				            <th scope="col">수정</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
                        do until rs.eof
							if rs("tax_yn") = "Y" then
								tax_view = "과세"
							  else
							  	tax_view = "비과세"
							end if
                        %>
				        <tr>
				          <td class="first"><a href="#" onClick="pop_Window('card_slip_list.asp?card_upjong=<%=rs("card_upjong")%>','card_slip_list_pop','scrollbars=yes,width=1000,height=300')"><%=rs("card_upjong")%></a></td>
				          <td><%=rs("account")%>&nbsp;</td>
				          <td><%=rs("account_item")%>&nbsp;</td>
				          <td><%=tax_view%>&nbsp;</td>
				          <td><a href="card_cust_upjong_mg.asp?card_upjong=<%=rs("card_upjong")%>&u_type=<%="U"%>">수정</a></td>
			            </tr>
				        <%
							rs.movenext()
						loop
						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="34%" valign="top"><form method="post" name="frm" action="card_cust_upjong_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				        <colgroup>
				          <col width="30%" >
				          <col width="*" >
			            </colgroup>
				          <tbody>
				            <tr>
				              <th>업종명</th>
				              <td class="left">
                    		<% if u_type = "U" then	%>
                              <%=card_upjong%>	
                              <input name="card_upjong" type="hidden" id="card_upjong" value="<%=card_upjong%>">
                            <%   else	%>
                              <input name="card_upjong" type="text" id="card_upjong" onKeyUp="checklength(this,30)" value="<%=card_upjong%>" notnull errname="업종명">
							<% end if	%>
                              </td>
			                </tr>
				            <tr>
				              <th>계정과목</th>
				              <td class="left">
                                <select name="account_view" id="account_view" style="width:200px">
                                    <option value="" <% if account_view = "" then %>selected<% end if %>>선택</option>
							<%
                                    Sql="select * from account_item where cost_yn = 'Y' or cost_yn = 'C' order by account_name, account_item asc"
                                    rs_acc.Open Sql, Dbconn, 1
                                    do until rs_acc.eof
										account = rs_acc("account_name") + "-" + rs_acc("account_item")
                            %>
                                    <option value='<%=account%>' <%If account_view = account then %>selected<% end if %>><%=account%></option>
                            <%
										rs_acc.movenext()
									loop
									rs_acc.close()						
                            %>
                                </select>
                              </td>
			                </tr>
				            <tr>
				              <th>과세구분</th>
				              <td class="left"><input type="radio" name="tax_yn" value="Y" <% if tax_yn = "Y" then %>checked<% end if %> style="width:40px" ID="Radio1">과세
				                <input type="radio" name="tax_yn" value="N" <% if tax_yn = "N" then %>checked<% end if %> style="width:40px" ID="Radio2">비과세 </td>
			                </tr>
			              </tbody>
			            </table>
						<br>
				        <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				        <input type="hidden" name="view_c" value="<%=view_c%>" ID="Hidden1">
				        <input type="hidden" name="field_view" value="<%=field_view%>" ID="Hidden1">
				        <div align=center>
                        	<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        </div>
			          </form></td>
			        </tr>
				    <tr>
				      <td width="49%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="49%">&nbsp;</td>
			        </tr>
			      </table>
                </div>
			</div>				
	</div>        				
	</body>
</html>

