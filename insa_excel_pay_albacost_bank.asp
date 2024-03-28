<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
rever_yymm=request("rever_yymm")
to_date=request("to_date")
view_bank=Request("view_bank")

curr_date = datevalue(mid(cstr(now()),1,10))

curr_yyyy = mid(cstr(rever_yymm),1,4)
curr_mm = mid(cstr(rever_yymm),6,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 사업소득 은행이체 내역(" + view_bank + ")"

savefilename = title_line +".xls"

	sum_alba_pay = 0
	sum_alba_trans = 0
	sum_alba_meals = 0
	sum_alba_other = 0
	sum_tax_amt1 = 0
	sum_tax_amt2 = 0
	sum_give_total = 0
	
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
Set Rs_alba = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_bank = "전체" then 
       Sql = "select * from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,bank_name,draft_no ASC"
   else
       Sql = "select * from pay_alba_cost where (rever_yymm = '"+rever_yymm+"' ) and (company = '"+view_condi+"') and (bank_name = '"+view_bank+"') ORDER BY company,bank_name,draft_no ASC"
end if
Rs.Open Sql, Dbconn, 1

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
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">등록번호</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">성명</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">등록일</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">회사</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">부서</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">소득구분</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">이체은행</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">계좌번호</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">예금주명</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">차인지급액</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">실지급액</div></td>
  </tr>
    <%
		do until rs.eof 
		
		  draft_no = rs("draft_no")
		  pay_count = pay_count + 1
				  
		  sum_alba_pay = sum_alba_pay + int(rs("alba_pay"))
	      sum_alba_trans = sum_alba_trans + int(rs("alba_trans"))
	      sum_alba_meals = sum_alba_meals + int(rs("alba_meals"))
	      sum_alba_other = sum_alba_other + int(rs("alba_other"))
	      sum_tax_amt1 = sum_tax_amt1 + int(rs("tax_amt1"))
	      sum_tax_amt2 = sum_tax_amt2 + int(rs("tax_amt2"))
          sum_give_total = sum_give_total + int(rs("alba_give_total"))
							  
	      curr_pay = int(rs("alba_give_total")) - (int(rs("tax_amt1")) + int(rs("tax_amt2")))
							  
		  Sql = "SELECT * FROM emp_alba_mst where draft_no = '"&draft_no&"'"
          Set rs_alba = DbConn.Execute(SQL)
		  if not rs_alba.eof then
		   		draft_date = rs_alba("draft_date")
	         else
	    		draft_date = ""
          end if
          rs_alba.close()

	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=rs("draft_no")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("draft_man")%></div></td>
    <td width="110"><div align="center" class="style1"><%=draft_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("company")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("org_name")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("draft_tax_id")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("bank_name")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("account_no")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("account_name")%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(curr_pay,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(curr_pay,0)%></div></td>
  </tr>
	<%
	    Rs.MoveNext()
	loop
	
	sum_curr_pay = sum_give_total - (sum_tax_amt1 + sum_tax_amt2)
	
	%>
    
  <tr>    
    <th colspan="9" style=" border-top:1px solid #e3e3e3;"><div align="center" class="style1">총계</div></th>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
  </tr>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
