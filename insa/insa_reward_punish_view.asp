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
Dim emp_name, title_line, rsPunish

emp_no = Request.QueryString("emp_no")
emp_name = Request.QueryString("emp_name")

title_line = "◈ 상벌 사항 ◈"

objBuilder.Append "SELECT app_date, app_id, app_id_type, app_reward, app_start_date, app_finish_date, "
objBuilder.Append "	app_comment, app_to_grade, app_to_position, app_to_company, app_to_org, app_to_orgcode "
objBuilder.Append "FROM emp_appoint "
objBuilder.Append "WHERE app_empno = '"&emp_no&"' "
objBuilder.Append "	AND (app_id = '포상발령' OR app_id = '징계발령') "
objBuilder.Append "ORDER BY app_empno, app_date, app_seq ASC "

Set rsPunish = DBConn.Execute(objBuilder.ToString())
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
							<col width="10%" >
							<col width="12%" >
							<col width="17%" >
                            <col width="*" >
                            <col width="33%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">상벌일자</th>
                                <th>상벌유형</th>
                                <th>징계기간</th>
                                <th>상벌내용</th>
                                <th>직급/직책 및 소속</th>
 							</tr>
						</thead>
						<tbody>
						<%
						If rsPunish.EOF Or rsPunish.BOF Then
							Response.Write "<tr><td colspan='5' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsPunish.EOF Or rsPunish.BOF
								v_cnt = v_cnt + 1

							 'task_memo = replace(rs("career_task"),chr(34),chr(39))
							 'view_memo = task_memo
							 'if len(task_memo) > 10 then
							 ' 	view_memo = mid(task_memo,1,10) + ".."
							 'end if

						%>
							<tr>
							  <td><%=rsPunish("app_date")%>&nbsp;</td>
                        <%If rsPunish("app_id") = "포상발령" Then %>
						      <td class="left">(포상)<%=rsPunish("app_id_type")%>&nbsp;</td>
                              <td class="left">&nbsp;</td>
                              <td class="left"><%=rsPunish("app_reward")%>&nbsp;</td>
                        <%ElseIf rsPunish("app_id") = "징계발령" Then %>
                              <td class="left">(징계)<%=rsPunish("app_id_type")%>&nbsp;</td>
                              <td class="left"><%=rsPunish("app_start_date")%>∼<%=rsPunish("app_finish_date")%>&nbsp;</td>
                              <td class="left"><%=rsPunish("app_comment")%>&nbsp;</td>
                        <%End If %>
                              <td class="left">
								<%=rsPunish("app_to_grade")%>-<%=rsPunish("app_to_position")%>(<%=rsPunish("app_to_company")%>&nbsp;<%=rsPunish("app_to_org")%>(<%=rsPunish("app_to_orgcode")%>)
							  </td>
							</tr>
							<%
								rsPunish.MoveNext()
							Loop
						End If
						rsPunish.close() : Set rsPunish = Nothing
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