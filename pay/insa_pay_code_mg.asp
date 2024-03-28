<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder
'===================================================
'### Request & Params
'===================================================
Dim etc_type, u_type, etc_code, rsPayCode
Dim etc_name, etc_group, group_name, used_sw, emp_tax_id
Dim title_line, etc_code_last

etc_type = f_Request("etc_type")
u_type = f_Request("u_type")
etc_code = f_Request("etc_code")

If etc_type = "" Then
	etc_type = "50"
End If



If u_type = "U" Then
	'sql = "select * from emp_etc_code where emp_etc_code = '" & etc_code & "'"
	objBuilder.Append "SELECT emp_etc_code, emp_etc_name, emp_etc_group, emp_group_name, "
	objBuilder.Append "	emp_used_sw, emp_tax_id "
	objBuilder.Append "FROM emp_etc_code "
	objBuilder.Append "WHERE emp_etc_code = '"&etc_code&"';"

	Set rs_etc = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	etc_code = rs_etc("emp_etc_code")
	etc_name = rs_etc("emp_etc_name")
	etc_group = rs_etc("emp_etc_group")
	group_name = rs_etc("emp_group_name")
	used_sw = rs_etc("emp_used_sw")
	emp_tax_id = rs_etc("emp_tax_id")

	rs_etc.Close() : Set rs_etc = Nothing
Else
	etc_code = ""
	etc_name = ""
	etc_group = ""
	group_name = ""
	used_sw = "Y"
	emp_tax_id = "9"
End If

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

			function frmsubmit(){
				document.condi_frm.submit();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				k = 0;

				for(j=0;j<2;j++){
					if(eval("document.frm.used_sw[" + j + "].checked")){
						k = k + 1
					}
				}

				if(k==0){
					alert ("이용여부를 선택하시기 바랍니다");
					return false;
				}

				var result = confirm('등록하시겠습니까?');

				if(result == true){
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
				<form action="/pay/insa_pay_code_mg.asp" method="post" name="condi_frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<strong>코드대분류 : </strong>
                                <select name="etc_type" id="etc_type" style="width:250px;">
                					<option value="">선택</option>
								<%
								Dim rs_type
								objBuilder.Append "SELECT emp_etc_seq, emp_etc_type, emp_type_name FROM emp_type_code ORDER BY emp_etc_type ASC;"

								Set rs_type = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								Do Until rs_type.EOF
									If rs_type("emp_etc_seq") = "1" Then
								%>
									<option value='<%=rs_type("emp_etc_type")%>' <%If rs_type("emp_etc_type") = etc_type Then %>selected<%End If %>><%=rs_type("emp_type_name")%></option>
								<%
									End If
									rs_type.MoveNext()
								Loop
								rs_type.Close() : Set rs_type = Nothing
								%>
            					</select>
                                <a href="#" onclick="javascript:frmsubmit();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
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
						Dim type_name, emp_tax_name

						etc_code_last = Int(etc_type&"01")

						'sql = "select * from emp_etc_code where emp_etc_type = '" & etc_type & "' order by emp_etc_code asc"
						objBuilder.Append "SELECT emp_type_name, emp_etc_code, emp_etc_type, emp_etc_name, emp_used_sw, "
						objBuilder.Append "	emp_tax_id "
						objBuilder.Append "FROM emp_etc_code "
						objBuilder.Append "WHERE emp_etc_type = '"&etc_type&"' "
						objBuilder.Append "ORDER BY emp_etc_code ASC;"

						Set rsPayCode = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

                        Do Until rsPayCode.EOF
                            type_name = rsPayCode("emp_type_name")
							emp_tax_name = ""
                        %>
				        <tr>
				          <td class="first"><%=rsPayCode("emp_etc_code")%></td>
				          <td>
							<a href="/insa_pay_code_mg.asp?etc_code=<%=rsPayCode("emp_etc_code")%>&etc_type=<%=rsPayCode("emp_etc_type")%>&u_type=U"><%=rsPayCode("emp_etc_name")%></a>
						  </td>
				          <td>&nbsp;<%=rsPayCode("emp_type_name")%></td>
				          <td>&nbsp;<%=rsPayCode("emp_used_sw")%></td>
				          <td>&nbsp;
						  <%
						  Select Case rsPayCode("emp_tax_id")
							Case "1"
								Response.Write "과세"
							Case "2"
								Response.Write "비과세"
							Case "3"
								Response.Write "감면세액"
							Case "9"
								Response.Write "해당없음"
						  End Select
						  %>
						  </td>
			            </tr>
				        <%
							etc_code_last = rsPayCode("emp_etc_code") + 1

							rsPayCode.movenext()
						Loop
						rsPayCode.Close() : Set rsPayCode = Nothing
						DBConn.Close() : Set DBConn = Nothing

						If etc_code_last < 1000 Then
							etc_code_last = "0"&CStr(etc_code_last)
						End If

						If u_type = "U" Then
							etc_code_last = etc_code
						End If
						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="49%" valign="top">
						<form method="post" name="frm" action="/pay/insa_pay_code_reg_ok.asp">
							<input name="etc_code" type="hidden" value="<%=etc_code_last%>"/>
							<input type="hidden" name="u_type" value="<%=u_type%>"/>
							<input type="hidden" name="etc_type" value="<%=etc_type%>"/>
							<input type="hidden" name="type_name" value="<%=type_name%>"/>

				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="25%">구분코드</th>
				              <td class="left"><%=etc_code_last%></td>
			                </tr>
				            <tr>
				              <th>코드명</th>
				              <td class="left">
								<input type="text" name="etc_name" id="etc_name" onKeyUp="checklength(this,20)" value="<%=etc_name%>" notnull errname="코드명"/>
							  </td>
			                </tr>
				            <tr>
				              <th>그룹코드</th>
				              <td class="left">
								<input type="text" name="etc_group" id="etc_group" size="2" maxlength="2" onlyposint value="<%=etc_group%>"/>
							  </td>
			                </tr>
				            <tr>
				              <th>그룹명</th>
				              <td class="left">
								<input type="text" name="group_name" id="group_name" size="22" onKeyUp="checklength(this, 20)" value="<%=group_name%>"/>
							  </td>
			                </tr>
				            <tr>
				              <th>사용여부</th>
				              <td class="left">
								<input type="radio" name="used_sw" value="Y" <%If used_sw = "Y" Then %>checked<%End If %> title="사용여부" style="width:40px;"/>
				                이용가능
				                <input type="radio" name="used_sw" value="N" <%If used_sw = "N" Then %>checked<%End If %> title="사용여부" style="width:40px;"/>
				                이용불가
							  </td>
			                </tr>
                            <tr>
                              <th>과세구분<br>급여관련만..</th>
                              <td class="left">
                                <select name="emp_tax_id" id="emp_type" value="<%=emp_tax_id%>" style="width:90px;">
			            	        <option value="" <%If emp_tax_id = "" Then %>selected<%End If %>>선택</option>
				                    <option value='1' <%If emp_tax_id = "1" Then %>selected<%End If %>>과세</option>
                                    <option value='2' <%If emp_tax_id = "2" Then %>selected<%End If %>>비과세</option>
				                    <option value='3' <%If emp_tax_id = "3" Then %>selected<%End If %>>감면소득</option>
                                    <option value='9' <%If emp_tax_id = "9" Then %>selected<%End If %>>해당없음</option>
                                </select>
                              </td>
                            </tr>
			              </tbody>
			            </table>
						<br/>
				        <div align="center">
                        	<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();"/></span>
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

