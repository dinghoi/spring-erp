<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")
in_pmg_id = request("pmg_id") 

curr_date = datevalue(mid(cstr(now()),1,10))

give_date = to_date '지급일

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)

if in_pmg_id = "2" then 
   pmg_id_name = "상여금" 
   elseif in_pmg_id = "3" then 
          pmg_id_name = "추천인인센티브" 
          elseif in_pmg_id = "4" then 
		         pmg_id_name = "연차수당" 
end if

title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + pmg_id_name +" 내역서(개인별)"

'당월 퇴사자 포함
st_es_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + "01"

savefilename = title_line +".xls"
'savefilename = "입사자 현황 -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"
'response.write(savefilename)

	sum_base_pay = 0
	sum_give_tot = 0

    sum_epi_amt = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if in_empno = "" then
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_no,emp_name ASC"
   else  
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_es_date&"') and (emp_company = '"+view_condi+"') and (emp_no = '"+in_empno+"') ORDER BY emp_company,emp_bonbu,emp_no,emp_name ASC"
end if
Rs_emp.Open Sql, Dbconn, 1

'Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
'Rs.Open Sql, Dbconn, 1

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
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="11" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">인적사항</div></td>
    <td style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">수당</div></td>
    <td colspan="5" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">공제 및 차인지급액</div></td>
  </tr>
  <tr>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">귀속년월</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">지급일</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사번</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성  명</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">입사일</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">직급</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">본부</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">사업부</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">팀</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">부서</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">지급액</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">고용보험</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">소득세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">지방소득세</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">공제합계</div></td>  
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">차인지급액</div></td>  
  </tr>
    <%
		do until Rs_emp.eof 
		
		  emp_no = Rs_emp("emp_no")
		  emp_name = Rs_emp("emp_name")
		  emp_grade = Rs_emp("emp_grade")
		  emp_position = Rs_emp("emp_position")
		  emp_in_date = rs_emp("emp_in_date")
		  
		  Sql = "SELECT * FROM pay_month_give where pmg_yymm = '"&pmg_yymm&"' and pmg_emp_no = '"&emp_no&"' and pmg_id = '"&in_pmg_id&"' and (pmg_company = '"+view_condi+"')"
          Set rs_give = DbConn.Execute(SQL)
		  if not rs_give.eof then
                   pmg_company = rs_give("pmg_company")
				   pmg_bonbu = rs_give("pmg_bonbu")
				   pmg_saupbu = rs_give("pmg_saupbu")
				   pmg_team = rs_give("pmg_team")
				   pmg_org_name = rs_give("pmg_org_name")
				   pmg_base_pay = rs_give("pmg_base_pay")
		  		   pmg_give_total = rs_give("pmg_give_total")
	         else
                   pmg_company = ""
				   pmg_bonbu = ""
				   pmg_saupbu = ""
				   pmg_team = ""
				   pmg_org_name = ""
				   pmg_base_pay = 0
			       pmg_give_total = 0
          end if
          rs_give.close()
		  
		  pay_count = pay_count + 1
					  
		  sum_base_pay = sum_base_pay + pmg_base_pay
	      sum_give_tot = sum_give_tot + pmg_give_total
		  
	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=pmg_yymm%></div></td>
    <td width="110"><div align="center" class="style1"><%=to_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_no%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_name%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_grade%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_company%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_bonbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_saupbu%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_team%></div></td>
    <td width="110"><div align="center" class="style1"><%=pmg_org_name%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_base_pay,0)%></div></td>
    <%
	      Sql = "SELECT * FROM pay_month_deduct where de_yymm = '"&pmg_yymm&"' and de_emp_no = '"&emp_no&"' and de_id = '"&in_pmg_id&"' and (de_company = '"+view_condi+"')"
          Set Rs_dct = DbConn.Execute(SQL)
		  if not Rs_dct.eof then
					de_epi_amt = Rs_dct("de_epi_amt")
					de_income_tax = Rs_dct("de_income_tax")
					de_wetax = Rs_dct("de_wetax")
					de_deduct_tot = Rs_dct("de_deduct_total")
	           else
                    de_deduct_tot = 0
					de_epi_amt = 0
					de_income_tax = 0
					de_wetax = 0
           end if
           Rs_dct.close()
		   
		   pmg_curr_pay = pmg_give_total - de_deduct_tot
							  
           sum_epi_amt = sum_epi_amt + de_epi_amt
           sum_income_tax = sum_income_tax + de_income_tax
           sum_wetax = sum_wetax + de_wetax
		   sum_deduct_tot = sum_deduct_tot + de_deduct_tot
							  
    %>    
    
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_epi_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_income_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_deduct_tot,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_curr_pay,0)%></div></td>
  </tr>
	<%
	    Rs_emp.MoveNext()
	loop
	
	sum_curr_pay = sum_give_tot - sum_deduct_tot
	
	%>
    
  <tr>    
    <th colspan="11" style=" border-top:1px solid #e3e3e3;"><div align="center" class="style1">총계</div></th>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_base_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_epi_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_income_tax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_wetax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_deduct_tot,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
  </tr>
</table>
</body>
</html>
<%
Rs_emp.Close()
Set Rs_emp = Nothing
%>
