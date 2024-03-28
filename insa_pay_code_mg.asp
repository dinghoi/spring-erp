<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim etc_code_last
dim e_etc_type
etc_type = request("etc_type")
u_type = request("u_type")
etc_code = request("etc_code")

if etc_type = "" then
	etc_type = request.form("etc_type")
end if

if etc_type = "" then
	etc_type = "50"
end if

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

sql = "select * from emp_etc_code where emp_etc_type = '" & etc_type & "' order by emp_etc_code asc"
Rs.Open Sql, Dbconn, 1

if u_type = "U" then
	sql = "select * from emp_etc_code where emp_etc_code = '" & etc_code & "'"
	Set rs_etc=DbConn.Execute(Sql)
	etc_code = rs_etc("emp_etc_code")
	etc_name = rs_etc("emp_etc_name")
	etc_group = rs_etc("emp_etc_group")
	group_name = rs_etc("emp_group_name")
	used_sw = rs_etc("emp_used_sw")
	emp_tax_id = rs_etc("emp_tax_id")
  else
	etc_code = ""
	etc_name = ""
	etc_group = ""
	group_name = ""
	used_sw = "Y"
	emp_tax_id = "9"
end if

title_line = "인사.급여 기본코드 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmsubmit () {
				document.condi_frm.submit ();
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
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
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_code_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa_pay_code_mg.asp" method="post" name="condi_frm">
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
								Sql="select emp_etc_seq, emp_etc_type, emp_type_name from emp_type_code order by emp_etc_type asc"
								Rs_type.Open Sql, Dbconn, 1
								do until Rs_type.eof
									if rs_type("emp_etc_seq") = "1" then
							%>
                					<option value='<%=Rs_type("emp_etc_type")%>' <% if Rs_type("emp_etc_type") = etc_type then %>selected<% end if %>><%=Rs_type("emp_type_name")%></option>
                			<%
									end if
									Rs_type.movenext()
								loop
								Rs_type.Close()
							%>
            					</select>
                                <a href="#" onclick="javascript:frmsubmit();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="49%" height="356"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="15%" >
				          <col width="*" >
				          <col width="15%" >
				          <col width="15%" >
				          <col width="25%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">구분코드</th>
				            <th scope="col">코드명</th>
				            <th scope="col">분류명</th>
                            <th scope="col">사용유무</th>
				            <th scope="col">과세여부</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
						etc_code_last = int(etc_type + "01")
'                        etc_code_last = etc_type + "01"
                        do until rs.eof
                            type_name = rs("emp_type_name")
							emp_tax_name = ""
                        %>
				        <tr>
				          <td class="first"><%=rs("emp_etc_code")%></td>
				          <td><a href="/insa_pay_code_mg.asp?etc_code=<%=rs("emp_etc_code")%>&etc_type=<%=rs("emp_etc_type")%>&u_type=<%="U"%>"><%=rs("emp_etc_name")%></a></td>
				          <td>&nbsp;<%=rs("emp_type_name")%></td>
				          <td>&nbsp;<%=rs("emp_used_sw")%></td>
                          <% If rs("emp_tax_id") = "1" then emp_tax_name = "과세" end if %>
                          <% If rs("emp_tax_id") = "2" then emp_tax_name = "비과세" end if %>
                          <% If rs("emp_tax_id") = "3" then emp_tax_name = "감면세액" end if %>
                          <% If rs("emp_tax_id") = "9" then emp_tax_name = "해당없음" end if %>
				          <td>&nbsp;<%=emp_tax_name%></td>
			            </tr>
				        <%
							etc_code_last = rs("emp_etc_code") + 1
						rs.movenext()
						loop
						if etc_code_last < 1000 then
							etc_code_last = "0" + cstr(etc_code_last)
						end if
						if u_type = "U" then
							etc_code_last = etc_code
						end if
						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="49%" valign="top"><form method="post" name="frm" action="/insa_pay_code_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="25%">구분코드</th>
				              <td class="left"><%=etc_code_last%><input name="etc_code" type="hidden" value="<%=etc_code_last%>"></td>
			                </tr>
				            <tr>
				              <th>코드명</th>
				              <td class="left"><input name="etc_name" type="text" id="etc_name" onKeyUp="checklength(this,20)" value="<%=etc_name%>" notnull errname="코드명"></td>
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
				              <th>사용여부</th>
				              <td class="left"><input type="radio" name="used_sw" value="Y" <% if used_sw = "Y" then %>checked<% end if %> title="사용여부" style="width:40px" ID="Radio1">
				                이용가능
				                <input type="radio" name="used_sw" value="N" <% if used_sw = "N" then %>checked<% end if %> title="사용여부" style="width:40px" ID="Radio2">
				                이용불가 </td>
			                </tr>
                            <tr>
                              <th>과세구분<br>급여관련만..</th>
                              <td class="left">
                                <select name="emp_tax_id" id="emp_type" value="<%=emp_tax_id%>" style="width:90px">
			            	        <option value="" <% if emp_tax_id = "" then %>selected<% end if %>>선택</option>
				                    <option value='1' <%If emp_tax_id = "1" then %>selected<% end if %>>과세</option>
                                    <option value='2' <%If emp_tax_id = "2" then %>selected<% end if %>>비과세</option>
				                    <option value='3' <%If emp_tax_id = "3" then %>selected<% end if %>>감면소득</option>
                                    <option value='9' <%If emp_tax_id = "9" then %>selected<% end if %>>해당없음</option>
                                </select>
                              </td>
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

