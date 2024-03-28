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
Dim from_date, to_date, view_condi, where_sql, main_title
Dim rs_emp, emp_org_baldate, emp_grade_date

from_date = Request.QueryString("from_date")
to_date = Request.QueryString("to_date")
view_condi = Request.QueryString("view_condi")

main_title = from_date & " ∼ " & to_date & " 입사자 현황"

If view_condi <> "전체" Then
	where_sql = "	AND eomt.org_company='" & view_condi & "' "
Else
	where_sql = ""
End If

objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_birthday, emtt.emp_grade, "
objBuilder.Append "	emtt.emp_job, emtt.emp_in_date, emtt.emp_org_name, emtt.emp_disab_grade, "
objBuilder.Append "	emtt.emp_last_edu, emtt.emp_position, emtt.emp_disabled, "
objBuilder.Append "	emtt.emp_reside_company, emtt.emp_org_baldate, emtt.emp_grade_date, "
objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (emtt.emp_in_date >= '" & from_date & "' AND emtt.emp_in_date <= '" & to_date & "') "
objBuilder.Append where_sql
objBuilder.Append "ORDER BY emtt.emp_no, emtt.emp_name ASC "

Set rs_emp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript">
            /*function printWindow(){
        //		viewOff("button");
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 25; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }*/

			//프린트 함수 신규 작성[허정호_20220204]
			var printArea;
			var initBody;

			function fnPrint(id){
				printArea = document.getElementById(id);

				window.onbeforeprint = beforePrint;
				window.onafterprint = afterPrint;

				window.print();
			}

			function beforePrint(){
				initBody = document.body.innerHTML;
				document.body.innerHTML = printArea.innerHTML;
			}

			function afterPrint(){
				document.body.innerHTML = initBody;
			}
        </script>
		<style type="text/css">
        <!--
    	    .style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	    .style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
            .style12C {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
            .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
			.style18L {font-size: 18px; font-family: "굴림체", "돋움체", Seoul; text-align: left; }
            .style16BC {font-size: 16px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; text-align: center; }
            .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
            .style14C {font-size: 14px; font-family: "굴림체", "돋움체", Seoul; text-align: center; }
            .style14BC {font-size: 14px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; text-align: center; }
            .style32BC {font-size: 32px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; text-align: center; }
        -->
        </style>
        <style media="print">
        .noprint     { display: none }
        </style>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div class="noprint">
			<p>
				<a href="#" onClick="fnPrint('print_pg');"><img src="/image/printer.jpg" width="39" height="36" border="0" alt="출력하기" /></a>
			</p>
		</div>
		<!--<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8"></object>-->
		<div id="print_pg">
		<table width="1020" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td colspan="3" align="center" class="style32BC"><%=main_title%></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
		</table>
		<table width="1020" border="1" cellspacing="0" cellpadding="0">
			<tr>
				<td width="5%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">사번</span></td>
				<td width="5%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">성  명</span></td>
				<td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">생년월일</span></td>
				<td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">직급</span></td>
				<td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">직위</span></td>
				<td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">직책</span></td>
				<td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">입사일</span></td>
				<td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">소속</span></td>
				<td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">최종학력</span></td>
				<td width="8%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">장애여부</span></td>
				<td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">상주처회사</span></td>
				<td width="28%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">조직</span></td>
			</tr>
			<%
			Do Until rs_emp.EOF
				If rs_emp("emp_org_baldate") = "1900-01-01" Then
				   emp_org_baldate = ""
				Else
				   emp_org_baldate = rs_emp("emp_org_baldate")
				End If

				If rs_emp("emp_grade_date") = "1900-01-01" Then
				   emp_grade_date = ""
				Else
				   emp_grade_date = rs_emp("emp_grade_date")
				End If
			%>
			<tr>
				<td width="5%" height="30" align="center"><span class="style12C"><%=rs_emp("emp_no")%></span></td>
				<td width="5%" height="30" align="center"><span class="style12C"><%=rs_emp("emp_name")%></span></td>
				<td width="6%" height="30" align="center">
					<span class="style12C"><%=rs_emp("emp_birthday")%></span>
				</td>
				<td width="6%" height="30" align="center"><span class="style12C"><%=rs_emp("emp_grade")%></span></td>
				<td width="6%" height="30" align="center"><span class="style12C"><%=rs_emp("emp_job")%></span></td>
				<td width="6%" height="30" align="center">
					<span class="style12C"><%=rs_emp("emp_position")%>&nbsp;</span>
				</td>
				<td width="6%" height="30" align="center">
					<span class="style12C"><%=rs_emp("emp_in_date")%></span>
				</td>
				<td width="9%" height="30" align="center"><span class="style12C"><%=rs_emp("org_name")%></span></td>
				<td width="6%" height="30" align="center">
					<span class="style12C"><%=rs_emp("emp_last_edu")%>&nbsp;</span>
				</td>
				<td width="8%" height="30" align="center">
					<span class="style12C"><%=rs_emp("emp_disabled")%>&nbsp;<%=rs_emp("emp_disab_grade")%></span>
				</td>
				<td width="9%" height="30" align="center">
					<span class="style12C"><%=rs_emp("emp_reside_company")%>&nbsp;</span>
				</td>
				<td width="28%" height="30" align="left">
					<span class="style12C">
					<%
					Response.Write rs_emp("org_company")

					If f_toString(rs_emp("org_bonbu"), "") <> "" Then
						Response.Write " - " & rs_emp("org_bonbu")
					End If

					If f_toString(rs_emp("org_team"), "") <> "" Then
						Response.Write " - " & rs_emp("org_team")
					End If
					%>
					</span>
				</td>
			</tr>
			<%
				rs_emp.MoveNext()
			Loop
			rs_emp.close() : Set rs_emp = Nothing
			DBConn.Close() : Set DBConn = Nothing
			%>
		</table>
		</div>
	</body>
</html>