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
Dim title_line
Dim arrTemp, rsCareer

emp_no = f_Request("emp_no")

title_line = "◈ 이전경력 사항 ◈"

objBuilder.Append "CALL USP_INSA_CAREER_INFO('"&emp_no&"')"
Call Rs_Open(rsCareer, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsCareer.EOF Then
	arrTemp = rsCareer.getRows()
End If

Call Rs_Close(rsCareer)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

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
        						<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:60px;" class="no-input" readonly="true"/>
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%Call EmpInfo_Name(emp_no)%>" style="width:100px;" class="no-input" readonly="true"/>
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="26%" >
                            <col width="18%" >
                            <col width="14%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">재직기간</th>
                                <th scope="col">회사명</th>
                                <th scope="col">부서</th>
                                <th scope="col">직위</th>
                                <th scope="col">담당업무</th>
 							</tr>
						</thead>
						<tbody>
						<%
						DBConn.Close() : Set DBConn = Nothing

						Dim i
						Dim career_task, career_join_date, career_end_date, career_office, career_dept
						Dim career_position, task_memo, view_memo

						If isArray(arrTemp) Then
							For i = 0 To UBound(arrTemp, 2)
								career_task = arrTemp(0, i)
								career_join_date = arrTemp(1, i)
								career_end_date = arrTemp(2, i)
								career_office = arrTemp(3, i)
								career_dept = arrTemp(4, i)
								career_position = arrTemp(5, i)

								task_memo = Replace(career_task, Chr(34), Chr(39))
								view_memo = task_memo

								If Len(task_memo) > 10 Then
							  		view_memo = Mid(task_memo, 1, 10) & "..."
								End If
						%>
							<tr>
								<td><%=career_join_date%> ∼ <%=career_end_date%>&nbsp;</td>
								<td><%=career_office%>&nbsp;</td>
                                <td><%=career_dept%>&nbsp;</td>
                                <td><%=career_position%>&nbsp;</td>
                                <td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
							</tr>
						<%
							Next
						Else
						%>
							<tr>
								<td class="first" colspan="5" style="height:30px;">조회된 내역이 없습니다.</td>
							</tr>
                        <%
						End If
						%>
						</tbody>
					</table>
				</div>
			</div>
		</div>
		<br>
		<div align="right">
			<a href="#" class="btnType04" onclick="close_win();" >닫기</a>&nbsp;&nbsp;
		</div>
		<br>
	</body>
</html>