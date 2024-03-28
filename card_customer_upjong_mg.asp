<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim etc_code_last
etc_type = request("etc_type")
u_type = request("u_type")
etc_code = request("etc_code")
mg_group = "0"
if etc_type = "" then
	etc_type = request.form("etc_type")
end if

if etc_type = "" then
	etc_type = "44"
end if

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")					
Set Rs_etc = Server.CreateObject("ADODB.Recordset")					
DbConn.Open dbconnect

sql = "select * from etc_code where etc_type = '" + etc_type + "' order by etc_code asc"
Rs.Open Sql, Dbconn, 1

if u_type = "U" then
	sql = "select * from etc_code where etc_code = '" + etc_code + "'"
	Set rs_etc=DbConn.Execute(Sql)
	etc_code = rs_etc("etc_code")
	etc_name = rs_etc("etc_name")
	etc_group = rs_etc("etc_group")
	group_name = rs_etc("group_name")
	mg_group = rs_etc("mg_group")
	etc_amt = rs_etc("etc_amt")
	used_sw = rs_etc("used_sw")
  else
	etc_name = ""
	etc_group = ""
	group_name = ""
	mg_group = "0"
	etc_amt = 0
	used_sw = "Y"
end if	

title_line = "회계 구분 코드 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
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
				if(document.frm.etc_name.value =="") {
					alert('업종명을 입력하세요');
					frm.etc_name.focus();
					return false;}
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.used_sw[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("이용여부를 선택하시기 바랍니다");
					return false;
				}	
			
				a=confirm('등록하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/account_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="account_etc_code_mg.asp" method="post" name="condi_frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<strong>코드대분류 : </strong>
                                <select name="etc_type" id="etc_type" style="width:250px">
                					<option>선택</option>
                            <%
								Sql="select * from type_code order by etc_type asc"
								Rs_type.Open Sql, Dbconn, 1
								do until Rs_type.eof
									if rs_type("etc_seq") = "3" then
							%>
                					<option value='<%=Rs_type("etc_type")%>' <% if Rs_type("etc_type") = etc_type then %>selected<% end if %>><%=Rs_type("type_name")%></option>
                			<%
									end if
									Rs_type.movenext()  
								loop 
								Rs_type.Close()
							%>
            					</select>
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
				          <col width="9%" >
				          <col width="*" >
				          <col width="9%" >
				          <col width="20%" >
				          <col width="12%" >
				          <col width="12%" >
				          <col width="10%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">구분코드</th>
				            <th scope="col">코드명</th>
				            <th scope="col">그룹코드</th>
				            <th scope="col">그룹명</th>
				            <th scope="col">관리그룹</th>
				            <th scope="col">금액</th>
				            <th scope="col">사용유무</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
                        etc_code_last = etc_type + "01"
                        do until rs.eof
                            type_name = rs("type_name")
							if rs("etc_amt") = "" or isnull(rs("etc_amt")) then
								etc_amt =0
							  else
							  	etc_amt = rs("etc_amt")
							end if
							if rs("mg_group") = "" or isnull(rs("mg_group")) or rs("mg_group") = "0" then
								mg_group_view = "공통"
							  else
							  	sql = "select * from type_code where etc_type = '91' and etc_seq ='"&rs("mg_group")&"'"
								Set rs_type=DbConn.Execute(Sql)
								mg_group_view = rs_type("type_name")
								rs_type.close()
							end if
                        %>
				        <tr>
				          <td class="first"><%=rs("etc_code")%></td>
				          <td><a href="account_etc_code_mg.asp?etc_code=<%=rs("etc_code")%>&etc_type=<%=rs("etc_type")%>&u_type=<%="U"%>"><%=rs("etc_name")%></a></td>
				          <td>&nbsp;<%=rs("etc_group")%></td>
				          <td>&nbsp;<%=rs("group_name")%></td>
				          <td><%=mg_group_view%></td>
				          <td><%=formatnumber(etc_amt,0)%></td>
				          <td><%=rs("used_sw")%></td>
			            </tr>
				        <%
							etc_code_last = rs("etc_code") + 1
						rs.movenext()
						loop
						if etc_code_last < 1000 then
							etc_code_last = "0" + cstr(etc_code_last)
						end if
						if u_type = "U" then
							etc_code_last = etc_code
						  else
						  	etc_amt = 0
						end if
						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="34%" valign="top"><form method="post" name="frm" action="account_etc_code_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="25%">구분코드</th>
				              <td class="left"><%=etc_code_last%><input name="etc_code" type="hidden" value="<%=etc_code_last%>"></td>
			                </tr>
				            <tr>
				              <th>코드명</th>
				              <td class="left"><input name="etc_name" type="text" id="etc_name" onKeyUp="checklength(this,20)" value="<%=etc_name%>"></td>
			                </tr>
				            <tr>
				              <th>그룹코드</th>
				              <td class="left"><input name="etc_group" type="text" id="etc_group" size="2" maxlength="2" onlyposint value="<%=etc_group%>"></td>
			                </tr>
				            <tr>
				              <th>그룹명</th>
				              <td class="left"><input name="group_name" type="text" id="group_name" size="22" onKeyUp="checklength(this,20)" value="<%=group_name%>"></td>
			                </tr>
				            <tr>
				              <th>금액</th>
				              <td class="left"><input name="etc_amt" type="text" id="etc_amt" value="<%=formatnumber(etc_amt,0)%>" onKeyUp="plusComma(this);" style="width:80px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>사용여부</th>
				              <td class="left"><input type="radio" name="used_sw" value="Y" <% if used_sw = "Y" then %>checked<% end if %> title="사용여부" style="width:40px" ID="Radio1">
				                이용가능
				                <input type="radio" name="used_sw" value="N" <% if used_sw = "N" then %>checked<% end if %> title="사용여부" style="width:40px" ID="Radio2">
				                이용불가 </td>
			                </tr>
			              </tbody>
			            </table>
						<br>
				        <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				        <input type="hidden" name="etc_type" value="<%=etc_type%>" ID="Hidden1">
				        <input type="hidden" name="type_name" value="<%=type_name%>" ID="Hidden1">
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

