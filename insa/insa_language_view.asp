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
Dim emp_name, rsLang, title_line

emp_no = Request("emp_no")
emp_name = Request("emp_name")

title_line = " 어학 능력 "

objBuilder.Append "SELECT lang_id, lang_id_type, lang_point, lang_grade, lang_get_date "
objBuilder.Append "FROM emp_language "
objBuilder.Append "WHERE lang_empno = '" & emp_no & "' "
objBuilder.Append "ORDER BY lang_empno, lang_seq ASC "

Set rsLang = DBConn.Execute(objBuilder.ToString())
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
							<col width="14%" >
							<col width="14%" >
                            <col width="10%" >
                            <col width="14%" >
                            <col width="14%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">어학구분</th>
                                <th scope="col">어학종류</th>
                                <th scope="col">점수</th>
                                <th scope="col">급수</th>
                                <th colspan="2" scope="col">취득일</th>
 							</tr>
						</thead>
						<tbody>
						<%
						If rsLang.EOF Or rsLang.BOF Then
							Response.Write "<tr><td colspan='5' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsLang.EOF
						%>
							<tr>
								<td><%=rsLang("lang_id")%>&nbsp;</td>
								<td><%=rsLang("lang_id_type")%>&nbsp;</td>
                                <td><%=rsLang("lang_point")%>&nbsp;</td>
                                <td><%=rsLang("lang_grade")%>&nbsp;</td>
                                <td colspan="2" class="left"><%=rsLang("lang_get_date")%>&nbsp;</td>
							</tr>
							<%
								rsLang.movenext()
							Loop
						End If
						rsLang.close() : Set rsLang = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
			</div>
			<br>
			<div align="right">
				<a href="#" class="btnType04" onclick="javascript:goAction();" >닫기</a>&nbsp;&nbsp;
			</div>
			<br>
	</body>
</html>