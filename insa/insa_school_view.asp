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
Dim arrTemp, rsSch

emp_no = f_Request("emp_no")

title_line = "◈ 학력 사항 ◈"

objBuilder.Append "CALL USP_INSA_SCHOOL_INFO('"&emp_no&"')"
Call Rs_Open(rsSch, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsSch.EOF Then
	arrTemp = rsSch.getRows()
End If

Call Rs_Close(rsSch)
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
							<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:60px;" class="no-input" readonly/>
						</label>
						<strong>성명 : </strong>
						<label>
							<input name="in_name" type="text" id="in_name" value="<%Call EmpInfo_Name(emp_no)%>" style="width:100px;" class="no-input" readonly/>
						</label>
						</p>
					</dd>
				</dl>
			</fieldset>
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="20%" >
						<col width="16%" >
						<col width="15%" >
						<col width="15%" >
						<col width="15%" >
						<col width="15%" >
						<col width="4%" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">기간</th>
							<th scope="col">학교명</th>
							<th scope="col">학과</th>
							<th scope="col">전공</th>
							<th scope="col">부전공</th>
							<th scope="col">학위구분</th>
							<th scope="col">졸업</th>
						</tr>
					</thead>
					<tbody>
					<%
					Dim i
					Dim sch_start_date, sch_end_date, sch_school_name, sch_dept, sch_major
					Dim sch_sub_major, sch_degree, sch_finish

					If IsArray(arrTemp) Then
						For i = 0 To UBound(arrTemp, 2)
							sch_start_date = arrTemp(0, i)
							sch_end_date = arrTemp(1, i)
							sch_school_name = arrTemp(2, i)
							sch_dept = arrTemp(3, i)
							sch_major = arrTemp(4, i)
							sch_sub_major = arrTemp(5, i)
							sch_degree = arrTemp(6, i)
							sch_finish = arrTemp(7, i)
					%>
						<tr>
							<td><%=sch_start_date%> ∼ <%=sch_end_date%>&nbsp;</td>
							<td><%=sch_school_name%>&nbsp;</td>
							<td><%=sch_dept%>&nbsp;</td>
							<td><%=sch_major%>&nbsp;</td>
							<td><%=sch_sub_major%>&nbsp;</td>
							<td><%=sch_degree%>&nbsp;</td>
							<td style=" border-bottom:1px solid #e3e3e3;"><%=sch_finish%>&nbsp;</td>
						</tr>
					<%
						Next
					Else
					%>
						<tr>
							<td class="first" colspan="6" style="height:30px;;">조회된 내역이 없습니다.</td>
						</tr>
					<%
					End If
					DBConn.Close() : Set DBConn = Nothing
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