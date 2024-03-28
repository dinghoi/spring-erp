<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### 작업 내역
'===================================================
' 허정호_20210721 :
'	- 신규 페이지 작성 및 코드 정리

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
Dim u_type, car_no, car_name, car_year, car_reg_date
Dim  owner_emp_name, owner_emp_no, use_date, title_line

u_type = request("u_type")
car_no = request("car_no")
car_name = request("car_name")
car_year = request("car_year")
car_reg_date = request("car_reg_date")
owner_emp_name = request("owner_emp_name")
owner_emp_no = request("owner_emp_no")

'나올수 없는 변수 값
'use_date = request("use_date")
'use_owner_em_no = request("use_owner_em_no")

use_date = ""
'필요없는 변구
'use_compamy = ""
'use_org_code = ""
'use_org_name = ""
'use_emp_name = ""
'use_emp_grade = ""
'use_end_date = ""

'view_condi = ""

title_line = "차량 운행자등록"

'나올수 없는 조건문
if u_type = "U" then

	sql = "select * from car_drive_user where use_car_no = '" + car_no + "' and use_owner_em_no = '" + use_owner_em_no + "' and use_date = '" + use_date + "'"
	set rs = dbconn.execute(sql)

    use_car_no = rs("ins_car_no")
	use_owner_emp_no = rs("use_owner_emp_no")
    use_date = rs("use_date")
    use_compamy = rs("use_compay")
    use_org_code = rs("use_org_code")
    use_org_name = rs("use_org_name")
    use_emp_name = rs("use_emp_name")
    use_emp_grade = rs("use_emp_grade")
    use_end_date = rs("use_end_date")
	rs.close()

	title_line = "차량 운행자변경"
end If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=use_date%>" );
			});

			function frmcheck(){
				//크롬에서 formcheck() 함수 사용시 자릿수 경고 체크로 미사용 처리[허정호_20210303]
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.use_date.value ==""){
					alert('변경일을 입력하세요');
					frm.use_date.focus();
					return false;
				}

				if(document.frm.emp_name.value ==""){
					alert('운행자를 입력하세요');
					frm.emp_name.focus();
					return false;
				}

				if(!confirm('입력하시겠습니까?')) return false;
				else return true;
			}
			/* 해당 항목 없음(cancel_col, info_col)[허정호_20210722]
			function update_view(){
				var c = document.frm.u_type.value;

				if (c == 'U'){
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}*/
        </script>
	</head>
	<!--<body onload="update_view()">-->
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_car_drvuser_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first" style="background:#FFFFE6">차량번호</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_no%>&nbsp;
                                <input name="car_no" type="hidden" value="<%=car_no%>" style="width:150px" readonly="true"></td>
								<th style="background:#FFFFE6">차종</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_name%>&nbsp;
                                <input name="car_name" type="hidden" value="<%=car_name%>" style="width:150px" readonly="true"></td>
							</tr>
                           	<tr>
								<th class="first" style="background:#FFFFE6">차량연식</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_year%>&nbsp;
                                <input name="car_year" type="hidden" value="<%=car_year%>" style="width:70px" readonly="true"></td>
                                <th style="background:#FFFFE6">차량등록일</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_reg_date%>&nbsp;
                                <input name="car_reg_date" type="hidden" value="<%=car_reg_date%>" style="width:70px" readonly="true"></td>
							</tr>
                            <tr>
								<th class="first" style="background:#FFFFE6">현 운행자</th>
								<td colspan="3" class="left" bgcolor="#FFFFE6"><%=owner_emp_name%>-<%=owner_emp_no%>&nbsp;

                                <input name="old_owner_emp_name" type="hidden" value="<%=owner_emp_name%>" style="width:70px" readonly="true">
                                <input name="old_owner_emp_no" type="hidden" value="<%=owner_emp_no%>" style="width:70px" readonly="true">
                                </td>
							</tr>
                            <tr>
								<th class="first">변경일</th>
								<td colspan="3" class="left"><input name="use_date" type="text" value="<%=use_date%>" style="width:70px" id="datepicker">
                                </td>
							</tr>
							<tr>
								<th class="first">운행자</th>
								<td colspan="3" class="left">

                                <input name="emp_name" type="text" id="emp_name" style="width:80px" value="<%'=use_owner_emp_name%>" readonly="true">
                                <input name="emp_grade" type="text" id="emp_grade" style="width:80px" value="<%'=use_emp_grade%>" readonly="true">
                                <input name="owner_emp_no" type="text" id="owner_emp_no" style="width:80px" value="<%'=use_owner_emp_no%>" readonly="true">

                                <a href="#" class="btnType03" onClick="pop_Window('/insa/insa_emp_select.asp?gubun=car&view_condi=<%'=view_condi%>','orgempselect','scrollbars=yes,width=600,height=400')">직원검색</a>
                                </td>
							</tr>
                            <tr>
								<th class="first">소속</th>
                                <td colspan="3" class="left">
                                <input name="emp_company" type="text" id="emp_company" style="width:120px" value="<%'=use_company%>" readonly="true">
                                <input name="emp_org_name" type="text" id="emp_org_name" style="width:120px" value="<%'=use_org_name%>" readonly="true">
                                <input name="emp_org_code" type="text" id="emp_org_code" style="width:120px" value="<%'=use_org_code%>" readonly="true">
                                </td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="저장" onclick="frmcheck();"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="toclose();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%'=u_type%>" ID="Hidden1">
			</form>
		</div>
	</body>
</html>

