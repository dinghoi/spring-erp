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
Dim rsQual, arrTemp

emp_no = f_Request("emp_no")

title_line = " 자격증 사항 "

objBuilder.Append "CALL USP_INSA_QUAL_INFO('"&emp_no&"')"
Call Rs_Open(rsQual, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rsQual.EOF Then
	arrTemp = rsQual.getRows()
End If

Call Rs_Close(rsQual)
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
						<col width="18%" >
						<col width="6%" >
						<col width="10%" >
						<col width="24%" >
						<col width="*" >
						<col width="16%" >
						<col width="8%" >
					</colgroup>
					<thead>
						<tr>
							<th class="first" scope="col">자격증&nbsp;종목</th>
							<th scope="col">등급</th>
							<th scope="col">합격년월일</th>
							<th scope="col">발급&nbsp;기관</th>
							<th scope="col">자격증&nbsp;&nbsp;번호</th>
							<th scope="col">경력수첩No.</th>
							<th scope="col">자격수당</th>
						</tr>
					</thead>
					<tbody>
					<%
					DBConn.Close : Set DBConn = Nothing

					Dim i, v_cnt
					Dim qual_pay_id, qual_type, qual_grade, qual_pass_date, qual_org
					Dim qual_no, qual_passport

					If isArray(arrTemp) Then
						For i = 0 To UBound(arrTemp, 2)
							qual_pay_id = arrTemp(0, i)
							qual_type = arrTemp(1, i)
							qual_grade = arrTemp(2, i)
							qual_pass_date = arrTemp(3, i)
							qual_org = arrTemp(4, i)
							qual_no = arrTemp(5, i)
							qual_passport = arrTemp(6, i)

							v_cnt = v_cnt + 1

							If f_toString(qual_pay_id, "") = "" Then
								qual_pay_id = "N"
							End If
					%>
						<tr>
							<td><%=qual_type%>&nbsp;</td>
							<td><%=qual_grade%>&nbsp;</td>
							<td><%=qual_pass_date%>&nbsp;</td>
							<td><%=qual_org%>&nbsp;</td>
							<td><%=qual_no%>&nbsp;</td>
							<td><%=qual_passport%>&nbsp;</td>
							<td style=" border-bottom:1px solid #e3e3e3;"><%=qual_pay_id%>&nbsp;</td>
						</tr>
					<%
						Next
					%>
						<tr>
							<td class="first" colspan="6"><%=v_cnt%>&nbsp;건이 조회되었습니다.</td>
						</tr>
					<%
					Else
					%>
						<tr>
							<td class="first" colspan="6" style="height:30px;">조회된 내역이 없습니다.</td>
						</tr>
					<%End If %>
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
</form>
</body>
</html>