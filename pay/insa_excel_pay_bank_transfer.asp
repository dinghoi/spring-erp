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
Dim view_condi, pmg_yymm, to_date, pmg_id, view_bank, give_date
Dim curr_yyyy, curr_mm, title_line, savefilename
Dim rsPay, sum_give_tot, sum_deduct_tot, sum_curr_pay
Dim pmg_emp_no, pmg_give_tot, emp_in_date, emp_jikmu, de_deduct_tot, pmg_curr_pay

view_condi = Request.QueryString("view_condi")
pmg_yymm = Request.QueryString("pmg_yymm")
to_date = Request.QueryString("to_date")
pmg_id = Request.QueryString("pmg_id")
view_bank = Request.QueryString("view_bank")

'curr_date = datevalue(mid(cstr(now()),1,10))

give_date = to_date '지급일

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여 은행이체 내역(" + view_bank + ")"

savefilename = title_line&".xls"
'savefilename = "입사자 현황 -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"

Call ViewExcelType(savefilename)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
	<tr bgcolor="#EFEFEF" class="style11">
		<td colspan="16" bgcolor="#FFFFFF"><div align="left" class="style2"><%=title_line%></div></td>
	</tr>
	<tr>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성명</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">입사일</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">직급</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">회사</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">부서</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">직무</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">이체은행</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">계좌번호</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">예금주명</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">차인지급액</div></td>
		<td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">실지급액</div></td>
	</tr>
    <%
	sum_give_tot = 0
	sum_deduct_tot = 0
	sum_curr_pay = 0

	'급여 정보 조회
	objBuilder.Append "SELECT pmgt.pmg_emp_no, pmgt.pmg_give_total, pmgt.pmg_emp_name, pmgt.pmg_grade, "
	objBuilder.Append "	pmgt.pmg_org_name, pmgt.pmg_bank_name, pmgt.pmg_account_no, pmgt.pmg_account_holder, "
	objBuilder.Append "	pmgt.pmg_company, "

	objBuilder.Append "	emtt.emp_in_date, emtt.emp_jikmu, "

	objBuilder.Append "	pmdt.de_nps_amt, pmdt.de_nhis_amt, pmdt.de_epi_amt, pmdt.de_longcare_amt, pmdt.de_income_tax, "
	objBuilder.Append "	pmdt.de_wetax, pmdt.de_year_incom_tax, pmdt.de_year_wetax, pmdt.de_year_incom_tax2, pmdt.de_year_wetax2, "
	objBuilder.Append "	pmdt.de_other_amt1, pmdt.de_sawo_amt, pmdt.de_hyubjo_amt, pmdt.de_school_amt, pmdt.de_nhis_bla_amt, "
	objBuilder.Append "	pmdt.de_long_bla_amt, pmdt.de_deduct_total "
	objBuilder.Append "FROM pay_month_give AS pmgt "
	objBuilder.Append "INNER JOIN emp_master AS emtt ON pmgt.pmg_emp_no = emtt.emp_no "
	objBuilder.Append "	AND (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' Or emtt.emp_end_date = '') "
	objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
	objBuilder.Append "	AND pmgt.pmg_company = pmdt.de_company "
	objBuilder.Append "	AND pmdt.de_id = '1' AND de_yymm = '"&pmg_yymm&"' "
	objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_id = '"&pmg_id&"' AND pmg_company = '"&view_condi&"' "

	If view_bank <> "전체" Then
		objBuilder.Append "AND pmgt.pmg_bank_name = '"&view_bank&"'"
	End If

	objBuilder.Append "ORDER BY pmgt.pmg_company, pmgt.pmg_bank_name, pmgt.pmg_emp_no ASC "

	Set rsPay = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	Do Until rsPay.EOF
		pmg_emp_no = rsPay("pmg_emp_no")
		pmg_give_tot = CLng(rsPay("pmg_give_total"))

		emp_in_date = rsPay("emp_in_date")
		emp_jikmu = rsPay("emp_jikmu")

		de_deduct_tot = CLng(rsPay("de_deduct_total"))

		sum_give_tot = sum_give_tot + pmg_give_tot
		sum_deduct_tot = sum_deduct_tot + de_deduct_tot

		pmg_curr_pay = pmg_give_tot - de_deduct_tot
	%>
	<tr valign="middle" class="style11">
		<td width="110"><div align="center" class="style1"><%=pmg_emp_no%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_emp_name")%></div></td>
		<td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_grade")%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_company")%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_org_name")%></div></td>
		<td width="110"><div align="center" class="style1"><%=emp_jikmu%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_bank_name")%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_account_no")%></div></td>
		<td width="110"><div align="center" class="style1"><%=rsPay("pmg_account_holder")%></div></td>
		<td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_curr_pay, 0)%></div></td>
		<td width="100"><div align="right" class="style1"><%=FormatNumber(pmg_curr_pay, 0)%></div></td>
		</tr>
		<%
			rsPay.MoveNext()
		Loop
		rsPay.Close() : Set rsPay = Nothing
		DBConn.Close() : Set DBConn = Nothing

		sum_curr_pay = sum_give_tot - sum_deduct_tot
		%>
	<tr>
		<th colspan="10" style=" border-top:1px solid #e3e3e3;"><div align="center" class="style1">총계</div></th>
		<td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=FormatNumber(sum_curr_pay, 0)%></div></td>
		<td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=FormatNumber(sum_curr_pay, 0)%></div></td>
	</tr>
</table>
</body>
</html>