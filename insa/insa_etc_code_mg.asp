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
Dim etc_type, u_type, etc_code
Dim etc_code_last, e_etc_type, title_line
Dim rsCode, etc_name, etc_group, group_name, used_sw, emp_tax_id
Dim type_name, rsEtc, rsEmp, rsType

etc_type = f_Request("etc_type")
u_type = f_Request("u_type")
etc_code = f_Request("etc_code")

If etc_type = "" Then
	etc_type = "01"
End If

title_line = "인사.급여 기본코드 관리"

objBuilder.Append "SELECT emp_type_name, emp_etc_code, emp_etc_type, emp_used_sw, emp_comment, emp_etc_name "
objBuilder.Append "FROM emp_etc_code WHERE emp_etc_type = '"&etc_type&"' ORDER BY emp_etc_code ASC "

Set rsCode = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If u_type = "U" Then
	objBuilder.Append "SELECT emp_etc_code, emp_etc_name, emp_etc_group, emp_group_name, emp_used_sw, emp_tax_id "
	objBuilder.Append "FROM emp_etc_code WHERE emp_etc_code = '"&etc_code&"' AND emp_etc_type = '"&etc_type&"' "

	Set rsEtc = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	etc_code = rsEtc("emp_etc_code")
	etc_name = rsEtc("emp_etc_name")
	etc_group = rsEtc("emp_etc_group")
	group_name = rsEtc("emp_group_name")

	If IsNull(group_name) Then
	   group_name = ""
	End If

	used_sw = rsEtc("emp_used_sw")
	emp_tax_id = rsEtc("emp_tax_id")

	rsEtc.Close() : Set rsEtc = Nothing
Else
	etc_code = ""
	etc_name = ""
	etc_group = ""
	group_name = ""
	used_sw = "Y"
	emp_tax_id = ""
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "5 1";
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

				var result = confirm('등록 하시겠습니까?');

				if(result){
					return true;
				}
				return false;
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_etc_code_mg.asp" method="post" name="condi_frm">
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
								objBuilder.Append "SELECT emp_etc_seq, emp_etc_type, emp_type_name "
								objBuilder.Append "FROM emp_type_code ORDER BY emp_etc_type ASC"

								Set rsType = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								Do Until rsType.EOF
									If rsType("emp_etc_seq") = "0" Then
							%>
                					<option value='<%=rsType("emp_etc_type")%>' <%If rsType("emp_etc_type") = etc_type Then %>selected<%End If %>><%=rsType("emp_type_name")%></option>
                			<%
									End If

									rsType.MoveNext()
								Loop
								rsType.Close() : Set rsType = Nothing
							%>
            					</select>
                                <a href="#" onclick="javascript:frmsubmit();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="49%" valign="top">
                      <table cellpadding="0" cellspacing="0" class="tableList">
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
				            <th scope="col">비고</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
						etc_code_last = Int(etc_type&"01")
'                        etc_code_last = etc_type + "01"

                        DO UNTIL rsCode.EOF
                            type_name = rsCode("emp_type_name")
                        %>
				        <tr>
				          <td class="first"><%=rsCode("emp_etc_code")%></td>
				          <td>
							<a href="/insa/insa_etc_code_mg.asp?etc_code=<%=rsCode("emp_etc_code")%>&etc_type=<%=rsCode("emp_etc_type")%>&u_type=U"><%=rsCode("emp_etc_name")%></a>
						  </td>
				          <td>&nbsp;<%=rsCode("emp_type_name")%></td>
				          <td>&nbsp;<%=rsCode("emp_used_sw")%></td>
				          <td>&nbsp;<%=rsCode("emp_comment")%></td>
			            </tr>
				        <%
							etc_code_last = rsCode("emp_etc_code") + 1

							rsCode.MoveNext()
						Loop
						rsCode.Close() : Set rsCode = Nothing

						If etc_code_last < 1000 Then
							etc_code_last = "0" & CStr(etc_code_last)
						End If

						If u_type = "U" Then
							etc_code_last = etc_code
						End If
						%>
			            </tbody>
			          </table>
                      </td>
                      </form>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="49%" valign="top">
						<form method="post" name="frm" action="/insa/insa_etc_code_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="25%">구분코드</th>
				              <td class="left"><%=etc_code_last%>
								<input name="etc_code" type="hidden" value="<%=etc_code_last%>">
							  </td>
			                </tr>
				            <tr>
				              <th>코드명</th>
				              <td class="left">
								<input name="etc_name" type="text" id="etc_name" onKeyUp="checklength(this,20)" value="<%=etc_name%>" notnull errname="코드명">
							  </td>
			                </tr>
				            <tr>
				              <th>그룹코드</th>
				              <td class="left">
								<input name="etc_group" type="text" id="etc_group" size="2" maxlength="2" onlyposint value="<%=etc_group%>">
							  </td>
			                </tr>
				            <tr>
							<%If etc_type = "13" Then %>
								<th>비용구분</th>
								<td class="left">
							<%
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '70' ORDER BY emp_etc_code ASC"

			  			        Set rsEmp = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							%>
								<select name="group_name" id="group_name" style="width:100px">
									<option value="" <%If group_name = "" Then %>selected<%End If %>>선택</option>
							<%
								Do Until rsEmp.EOF
			  				%>
                					<option value='<%=rsEmp("emp_etc_name")%>' <%If group_name = rsEmp("emp_etc_name") Then %>selected<%End If %>><%=rsEmp("emp_etc_name")%></option>
							<%
									rsEmp.MoveNext()
								Loop
								rsEmp.Close() : Set rsEmp = Nothing
							%>
                			     </select>
                              </td>
							<%Else %>
                              <th>그룹명</th>
				              <td class="left">
								<input name="group_name" type="text" id="group_name" size="22" onKeyUp="checklength(this,20)" value="<%=group_name%>">
							  </td>
							<%
							End If
							DBConn.Close() : Set DBConn = Nothing
							%>
			                </tr>
				            <tr>
								<th>사용여부</th>
								<td class="left">
									<input type="radio" name="used_sw" value="Y" <%If used_sw = "Y" Then %>checked<%End If %> title="사용여부" style="width:40px" />
									이용가능
									<input type="radio" name="used_sw" value="N" <%If used_sw = "N" Then %>checked<%End If %> title="사용여부" style="width:40px" />
									이용불가
								</td>
			                </tr>
			              </tbody>
			            </table>
						<br>
				        <input type="hidden" name="u_type" value="<%=u_type%>" />
				        <input type="hidden" name="etc_type" value="<%=etc_type%>" />
				        <input type="hidden" name="type_name" value="<%=type_name%>" />
				        <div align="center">
                        	<span class="btnType01">
								<input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1">
							</span>
                        </div>
			          </form>
					  </td>
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