<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim emp_name, title_line, rsComment

emp_no = Request.QueryString("emp_no")
emp_name = Request.QueryString("emp_name")

title_line = " 특이사항 "

objBuilder.Append "SELECT cmt_date, cmt_comment, cmt_org_name, cmt_org_code, cmt_company, cmt_bonbu, cmt_saupbu, cmt_team "
objBuilder.Append "FROM emp_comment "
objBuilder.Append "WHERE cmt_empno = '" & emp_no & "' "
objBuilder.Append "ORDER BY cmt_empno, cmt_date DESC "

Set rsComment = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction(){
			   window.close();
			}
		</script>
		<style type="text/css">
		.no-input{
			color:gray;
			background-color:#E0E0E0;
			border:1px solid #999999;
		}
		</style>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>사번 : </strong>
								<label>
        						<input type="text" name="in_empno" id="in_empno" value="<%=emp_no%>" style="width:60px; text-align:left;" class="no-input" readonly/>
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input type="text" name="in_name" id="in_name" value="<%=emp_name%>" style="width:100px; text-align:left;" class="no-input" readonly/>
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%">
							<col width="*">
                            <col width="14%">
                            <col width="40%">
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">발생일</th>
                                <th scope="col">특이사항</th>
                                <th scope="col">소속</th>
                                <th scope="col">조직</th>
 							</tr>
						</thead>
						<tbody>
						<%
						If rsComment.EOF Or rsComment.BOF Then
							Response.Write "<tr><td colspan='4' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsComment.EOF
						%>
							<tr>
								<td><%=rsComment("cmt_date")%>&nbsp;</td>
								<td class="left"><%=rsComment("cmt_comment")%>&nbsp;</td>
                                <td><%=rsComment("cmt_org_name")%>(<%=rsComment("cmt_org_code")%>)&nbsp;</td>
                                <td class="left"><%=rsComment("cmt_company")%>-<%=rsComment("cmt_bonbu")%>-<%=rsComment("cmt_saupbu")%>-<%=rsComment("cmt_team")%>&nbsp;</td>
							</tr>
						<%
								rsComment.movenext()
							Loop
						End If
						rsComment.close() : Set rsComment = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
			</div>
			<br>
			<div align="right">
				<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
			</div>
			<br>
	</body>
</html>