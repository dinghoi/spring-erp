<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

give_date = to_date '������

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " �޿� ������(" + view_condi + ")"

savefilename = title_line +".xls"
'savefilename = "�Ի��� ��Ȳ -- "+ to_date +""+ view_condi +"" + cstr(curr_date) + ".xls"
'response.write(savefilename)

	sum_base_pay = 0 
	sum_meals_pay = 0
	sum_postage_pay = 0
	sum_re_pay = 0
	sum_overtime_pay = 0
	sum_car_pay = 0
	sum_position_pay = 0
	sum_custom_pay = 0
	sum_job_pay = 0
	sum_job_support = 0
	sum_jisa_pay = 0
	sum_long_pay = 0
	sum_disabled_pay = 0
	sum_family_pay = 0
	sum_school_pay = 0
	sum_qual_pay = 0
	sum_other_pay1 = 0
	sum_other_pay2 = 0
	sum_other_pay3 = 0
	sum_tax_yes = 0
	sum_tax_no = 0
	sum_tax_reduced = 0
	sum_give_tot = 0
    sum_nps_amt = 0
    sum_nhis_amt = 0
    sum_epi_amt = 0
    sum_longcare_amt = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_year_incom_tax2 = 0
    sum_year_wetax2 = 0
    sum_other_amt1 = 0
    sum_sawo_amt = 0
    sum_hyubjo_amt = 0
    sum_school_amt = 0
    sum_nhis_bla_amt = 0
    sum_long_bla_amt = 0
	sum_deduct_tot = 0
	
	pay_count = 0	
	sum_curr_pay = 0	
' �����	
	org_base_pay = 0 
	org_meals_pay = 0
	org_postage_pay = 0
	org_re_pay = 0
	org_overtime_pay = 0
	org_car_pay = 0
	org_position_pay = 0
	org_custom_pay = 0
	org_job_pay = 0
	org_job_support = 0
	org_jisa_pay = 0
	org_long_pay = 0
	org_disabled_pay = 0
	org_family_pay = 0
	org_school_pay = 0
	org_qual_pay = 0
	org_other_pay1 = 0
	org_other_pay2 = 0
	org_other_pay3 = 0
	org_tax_yes = 0
	org_tax_no = 0
	org_tax_reduced = 0
	org_give_tot = 0
    org_nps_amt = 0
    org_nhis_amt = 0
    org_epi_amt = 0
    org_longcare_amt = 0
    org_income_tax = 0
    org_wetax = 0
	org_year_incom_tax = 0
    org_year_wetax = 0
	org_year_incom_tax2 = 0
    org_year_wetax2 = 0
    org_other_amt1 = 0
    org_sawo_amt = 0
    org_hyubjo_amt = 0
    org_school_amt = 0
    org_nhis_bla_amt = 0
    org_long_bla_amt = 0
	org_deduct_tot = 0
	
	org_pay_count = 0	
	org_curr_pay = 0
	
' ��	
	team_base_pay = 0 
	team_meals_pay = 0
	team_postage_pay = 0
	team_re_pay = 0
	team_overtime_pay = 0
	team_car_pay = 0
	team_position_pay = 0
	team_custom_pay = 0
	team_job_pay = 0
	team_job_support = 0
	team_jisa_pay = 0
	team_long_pay = 0
	team_disabled_pay = 0
	team_family_pay = 0
	team_school_pay = 0
	team_qual_pay = 0
	team_other_pay1 = 0
	team_other_pay2 = 0
	team_other_pay3 = 0
	team_tax_yes = 0
	team_tax_no = 0
	team_tax_reduced = 0
	team_give_tot = 0
    team_nps_amt = 0
    team_nhis_amt = 0
    team_epi_amt = 0
    team_longcare_amt = 0
    team_income_tax = 0
    team_wetax = 0
	team_year_incom_tax = 0
    team_year_wetax = 0
	team_year_incom_tax2 = 0
    team_year_wetax2 = 0
    team_other_amt1 = 0
    team_sawo_amt = 0
    team_hyubjo_amt = 0
    team_school_amt = 0
    team_nhis_bla_amt = 0
    team_long_bla_amt = 0
	team_deduct_tot = 0
	
	team_pay_count = 0	
	team_curr_pay = 0	

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
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

if view_condi = "��ü" then
      Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') ORDER BY pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_code,pmg_emp_no ASC"
   else
      Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_bonbu,pmg_saupbu,pmg_team,pmg_org_code,pmg_emp_no ASC"
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
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="9" style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��������</div></td>
    <td colspan="14" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><div align="center" class="style1">�⺻�޿� �� ������</div></td>
    <td colspan="14" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><div align="center" class="style1">���� �� �������޾�</div></td>
  </tr>
  <tr>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��  ��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�Ի���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">ȸ��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="center" class="style1">�μ�</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�⺻��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�Ĵ�</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��ź�</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ұޱ޿�</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����ٷμ���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��å����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">������ٹ���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ټӼ���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����μ���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�����հ�</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���ο���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ǰ�����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��뺸��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����纸���</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ҵ漼</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����ҵ漼</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��������ҵ漼</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�����������漼</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����������ҵ漼</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�������������漼</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">��Ÿ����</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���ȸ ȸ��</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">���ڱݻ�ȯ</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�ǰ����������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">����纸�������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">������</div></td>
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�����հ�</div></td>  
    <td style=" border-bottom:1px solid #e3e3e3;"><div align="right" class="style1">�������޾�</div></td>  
  </tr>
    <%
       if rs.eof or rs.bof then
		       bi_org = ""
			   bi_team = ""
		  else						  
			   if isnull(rs("pmg_saupbu")) or rs("pmg_saupbu") = "" then	
				 	   bi_org = ""
				  else
					   bi_org = rs("pmg_saupbu")
			   end if
			   if isnull(rs("pmg_team")) or rs("pmg_team") = "" then	
				 	   bi_team = ""
				  else
					   bi_team = rs("pmg_team")
			   end if
		end if		
		
		do until rs.eof 

          if isnull(rs("pmg_saupbu")) or rs("pmg_saupbu") = "" then
				    pmg_saupbu = ""
			 else
			        pmg_saupbu = rs("pmg_saupbu")
		  end if
		  if isnull(rs("pmg_team")) or rs("pmg_team") = "" then
		  	        pmg_team = ""
	 	     else
			        pmg_team = rs("pmg_team")
		  end if		

          if bi_team <> pmg_team then
		             team_curr_pay = team_give_tot - team_deduct_tot
	%>
                 <tr>
                    <td colspan="8" bgcolor="#EEFFFF" align="center"><%=bi_team%>&nbsp;&nbsp;&nbsp;����</div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_pay_count,0)%>&nbsp;��</td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_base_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_meals_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_postage_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_re_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_overtime_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_car_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_position_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_custom_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_job_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_job_support,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_jisa_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_long_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_disabled_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_give_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_nps_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_nhis_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_epi_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_longcare_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_income_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_incom_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_incom_tax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_wetax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_other_amt1,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_sawo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_school_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_nhis_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_long_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_hyubjo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_deduct_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_curr_pay,0)%></div></td>
                 </tr>
    <%
				 team_base_pay = 0 
				 team_meals_pay = 0
	             team_postage_pay = 0
	             team_re_pay = 0
	             team_overtime_pay = 0
	             team_car_pay = 0
	             team_position_pay = 0
	             team_custom_pay = 0
	             team_job_pay = 0
	             team_job_support = 0
              	 team_jisa_pay = 0
	             team_long_pay = 0
	             team_disabled_pay = 0
	             team_family_pay = 0
	             team_school_pay = 0
	             team_qual_pay = 0
	             team_other_pay1 = 0
	             team_other_pay2 = 0
	             team_other_pay3 = 0
	             team_tax_yes = 0
	             team_tax_no = 0
	             team_tax_reduced = 0
	             team_give_tot = 0
                 team_nps_amt = 0
                 team_nhis_amt = 0
                 team_epi_amt = 0
                 team_longcare_amt = 0
                 team_income_tax = 0
                 team_wetax = 0
	             team_year_incom_tax = 0
                 team_year_wetax = 0
				 team_year_incom_tax2 = 0
                 team_year_wetax2 = 0
                 team_other_amt1 = 0
                 team_sawo_amt = 0
                 team_hyubjo_amt = 0
                 team_school_amt = 0
                 team_nhis_bla_amt = 0
                 team_long_bla_amt = 0
	             team_deduct_tot = 0
	
	             team_pay_count = 0	
	             team_curr_pay = 0
				 
				 bi_team = pmg_team
		   end if

          if bi_org <> pmg_saupbu then
		            org_curr_pay = org_give_tot - org_deduct_tot
	%>
                 <tr>
				    <td colspan="8" bgcolor="#EEFFFF" align="center"><%=bi_org%>&nbsp;&nbsp;&nbsp;����ΰ�</div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_pay_count,0)%>&nbsp;��</td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_base_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_meals_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_postage_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_re_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_overtime_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_car_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_position_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_custom_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_job_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_job_support,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_jisa_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_long_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_disabled_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_give_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_nps_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_nhis_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_epi_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_longcare_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_income_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_incom_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_incom_tax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_wetax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_other_amt1,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_sawo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_school_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_nhis_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_long_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_hyubjo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_deduct_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_curr_pay,0)%></div></td>
                 </tr>
    <%
				 team_base_pay = 0 
				 team_meals_pay = 0
	             team_postage_pay = 0
	             team_re_pay = 0
	             team_overtime_pay = 0
	             team_car_pay = 0
	             team_position_pay = 0
	             team_custom_pay = 0
	             team_job_pay = 0
	             team_job_support = 0
              	 team_jisa_pay = 0
	             team_long_pay = 0
	             team_disabled_pay = 0
	             team_family_pay = 0
	             team_school_pay = 0
	             team_qual_pay = 0
	             team_other_pay1 = 0
	             team_other_pay2 = 0
	             team_other_pay3 = 0
	             team_tax_yes = 0
	             team_tax_no = 0
	             team_tax_reduced = 0
	             team_give_tot = 0
                 team_nps_amt = 0
                 team_nhis_amt = 0
                 team_epi_amt = 0
                 team_longcare_amt = 0
                 team_income_tax = 0
                 team_wetax = 0
	             team_year_incom_tax = 0
                 team_year_wetax = 0
				 team_year_incom_tax2 = 0
                 team_year_wetax2 = 0
                 team_other_amt1 = 0
                 team_sawo_amt = 0
                 team_hyubjo_amt = 0
                 team_school_amt = 0
                 team_nhis_bla_amt = 0
                 team_long_bla_amt = 0
	             team_deduct_tot = 0
	
	             team_pay_count = 0	
	             team_curr_pay = 0
				 
				 org_base_pay = 0 
				 org_meals_pay = 0
	             org_postage_pay = 0
	             org_re_pay = 0
	             org_overtime_pay = 0
	             org_car_pay = 0
	             org_position_pay = 0
	             org_custom_pay = 0
	             org_job_pay = 0
	             org_job_support = 0
              	 org_jisa_pay = 0
	             org_long_pay = 0
	             org_disabled_pay = 0
	             org_family_pay = 0
	             org_school_pay = 0
	             org_qual_pay = 0
	             org_other_pay1 = 0
	             org_other_pay2 = 0
	             org_other_pay3 = 0
	             org_tax_yes = 0
	             org_tax_no = 0
	             org_tax_reduced = 0
	             org_give_tot = 0
                 org_nps_amt = 0
                 org_nhis_amt = 0
                 org_epi_amt = 0
                 org_longcare_amt = 0
                 org_income_tax = 0
                 org_wetax = 0
	             org_year_incom_tax = 0
                 org_year_wetax = 0
				 org_year_incom_tax2 = 0
                 org_year_wetax2 = 0
                 org_other_amt1 = 0
                 org_sawo_amt = 0
                 org_hyubjo_amt = 0
                 org_school_amt = 0
                 org_nhis_bla_amt = 0
                 org_long_bla_amt = 0
	             org_deduct_tot = 0
	
	             org_pay_count = 0	
	             org_curr_pay = 0
				 
				 bi_org = pmg_saupbu
		   end if
		  
		  emp_no = rs("pmg_emp_no")
		  pmg_company = rs("pmg_company")
		  pmg_give_tot = rs("pmg_give_total")
		  
		  pay_count = pay_count + 1
		  sum_base_pay = sum_base_pay + int(rs("pmg_base_pay"))
	      sum_meals_pay = sum_meals_pay + int(rs("pmg_meals_pay"))
	      sum_postage_pay = sum_postage_pay + int(rs("pmg_postage_pay"))
	      sum_re_pay = sum_re_pay + int(rs("pmg_re_pay"))
	      sum_overtime_pay = sum_overtime_pay + int(rs("pmg_overtime_pay"))
	      sum_car_pay = sum_car_pay + int(rs("pmg_car_pay"))
          sum_position_pay = sum_position_pay + int(rs("pmg_position_pay"))
	      sum_custom_pay = sum_custom_pay + int(rs("pmg_custom_pay"))
	      sum_job_pay = sum_job_pay + int(rs("pmg_job_pay"))
	      sum_job_support = sum_job_support + int(rs("pmg_job_support"))
	      sum_jisa_pay = sum_jisa_pay + int(rs("pmg_jisa_pay"))
	      sum_long_pay = sum_long_pay + int(rs("pmg_long_pay"))
	      sum_disabled_pay = sum_disabled_pay + int(rs("pmg_disabled_pay"))
	      sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))
		  
		  org_pay_count = org_pay_count + 1
		  org_base_pay = org_base_pay + int(rs("pmg_base_pay"))
	      org_meals_pay = org_meals_pay + int(rs("pmg_meals_pay"))
	      org_postage_pay = org_postage_pay + int(rs("pmg_postage_pay"))
	      org_re_pay = org_re_pay + int(rs("pmg_re_pay"))
	      org_overtime_pay = org_overtime_pay + int(rs("pmg_overtime_pay"))
	      org_car_pay = org_car_pay + int(rs("pmg_car_pay"))
          org_position_pay = org_position_pay + int(rs("pmg_position_pay"))
	      org_custom_pay = org_custom_pay + int(rs("pmg_custom_pay"))
	      org_job_pay = org_job_pay + int(rs("pmg_job_pay"))
	      org_job_support = org_job_support + int(rs("pmg_job_support"))
	      org_jisa_pay = org_jisa_pay + int(rs("pmg_jisa_pay"))
	      org_long_pay = org_long_pay + int(rs("pmg_long_pay"))
	      org_disabled_pay = org_disabled_pay + int(rs("pmg_disabled_pay"))
	      org_give_tot = org_give_tot + int(rs("pmg_give_total"))
		  
		  team_pay_count = team_pay_count + 1
		  team_base_pay = team_base_pay + int(rs("pmg_base_pay"))
	      team_meals_pay = team_meals_pay + int(rs("pmg_meals_pay"))
	      team_postage_pay = team_postage_pay + int(rs("pmg_postage_pay"))
	      team_re_pay = team_re_pay + int(rs("pmg_re_pay"))
	      team_overtime_pay = team_overtime_pay + int(rs("pmg_overtime_pay"))
	      team_car_pay = team_car_pay + int(rs("pmg_car_pay"))
          team_position_pay = team_position_pay + int(rs("pmg_position_pay"))
	      team_custom_pay = team_custom_pay + int(rs("pmg_custom_pay"))
	      team_job_pay = team_job_pay + int(rs("pmg_job_pay"))
	      team_job_support = team_job_support + int(rs("pmg_job_support"))
	      team_jisa_pay = team_jisa_pay + int(rs("pmg_jisa_pay"))
	      team_long_pay = team_long_pay + int(rs("pmg_long_pay"))
	      team_disabled_pay = team_disabled_pay + int(rs("pmg_disabled_pay"))
	      team_give_tot = team_give_tot + int(rs("pmg_give_total"))
		  
		  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
          Set rs_emp = DbConn.Execute(SQL)
		  if not rs_emp.eof then
				emp_in_date = rs_emp("emp_in_date")
	         else
				emp_in_date = ""
          end if
          rs_emp.close()

	%>
  <tr valign="middle" class="style11">
    <td width="110"><div align="center" class="style1"><%=rs("pmg_emp_no")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_emp_name")%></div></td>
    <td width="110"><div align="center" class="style1"><%=emp_in_date%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_grade")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_company")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_bonbu")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_saupbu")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_team")%></div></td>
    <td width="110"><div align="center" class="style1"><%=rs("pmg_org_name")%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_base_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_meals_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_postage_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_re_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_overtime_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_car_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_position_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_custom_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_job_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_job_support"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_jisa_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_long_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_disabled_pay"),0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(rs("pmg_give_total"),0)%></div></td>
    <%
	      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"')  and (de_company = '"+pmg_company+"')"
          Set Rs_dct = DbConn.Execute(SQL)
		  if not Rs_dct.eof then
				de_nps_amt = int(Rs_dct("de_nps_amt"))
                de_nhis_amt = int(Rs_dct("de_nhis_amt"))
                de_epi_amt = int(Rs_dct("de_epi_amt"))
                de_longcare_amt = int(Rs_dct("de_longcare_amt"))
                de_income_tax = int(Rs_dct("de_income_tax"))
                de_wetax = int(Rs_dct("de_wetax"))
				de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
                de_year_wetax = int(Rs_dct("de_year_wetax"))
				de_year_incom_tax2 = int(Rs_dct("de_year_incom_tax2"))
                de_year_wetax2 = int(Rs_dct("de_year_wetax2"))
                de_other_amt1 = int(Rs_dct("de_other_amt1"))
                de_sawo_amt = int(Rs_dct("de_sawo_amt"))
                de_hyubjo_amt = int(Rs_dct("de_hyubjo_amt"))
                de_school_amt = int(Rs_dct("de_school_amt"))
                de_nhis_bla_amt = int(Rs_dct("de_nhis_bla_amt"))
                de_long_bla_amt = int(Rs_dct("de_long_bla_amt"))	
                de_deduct_tot = int(Rs_dct("de_deduct_total"))	
             else
				de_nps_amt = 0
                de_nhis_amt = 0
                de_epi_amt = 0
                de_longcare_amt = 0
                de_income_tax = 0
                de_wetax = 0
				de_year_incom_tax = 0
                de_year_wetax = 0
				de_year_incom_tax2 = 0
                de_year_wetax2 = 0
                de_other_amt1 = 0
                de_sawo_amt = 0
                de_hyubjo_amt = 0
                de_school_amt = 0
                de_nhis_bla_amt = 0
                de_long_bla_amt = 0
                de_deduct_tot = 0
           end if
           Rs_dct.close()
		   
		   pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  
	 	   sum_nps_amt = sum_nps_amt + de_nps_amt
           sum_nhis_amt = sum_nhis_amt + de_nhis_amt
           sum_epi_amt = sum_epi_amt + de_epi_amt
		   sum_longcare_amt = sum_longcare_amt + de_longcare_amt
           sum_income_tax = sum_income_tax + de_income_tax
           sum_wetax = sum_wetax + de_wetax
		   sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
           sum_year_wetax = sum_year_wetax + de_year_wetax
		   sum_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
           sum_year_wetax2 = sum_year_wetax2 + de_year_wetax2
           sum_other_amt1 = sum_other_amt1 + de_other_amt1
           sum_sawo_amt = sum_sawo_amt + de_sawo_amt
           sum_hyubjo_amt = sum_hyubjo_amt + de_hyubjo_amt
           sum_school_amt = sum_school_amt + de_school_amt
           sum_nhis_bla_amt = sum_nhis_bla_amt + de_nhis_bla_amt
           sum_long_bla_amt = sum_long_bla_amt + de_long_bla_amt
		   sum_deduct_tot = sum_deduct_tot + de_deduct_tot
		   
		   org_nps_amt = org_nps_amt + de_nps_amt
           org_nhis_amt = org_nhis_amt + de_nhis_amt
           org_epi_amt = org_epi_amt + de_epi_amt
		   org_longcare_amt = org_longcare_amt + de_longcare_amt
           org_income_tax = org_income_tax + de_income_tax
           org_wetax = org_wetax + de_wetax
		   org_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
           org_year_wetax = sum_year_wetax + de_year_wetax
		   org_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
           org_year_wetax2 = sum_year_wetax2 + de_year_wetax2
           org_other_amt1 = org_other_amt1 + de_other_amt1
           org_sawo_amt = org_sawo_amt + de_sawo_amt
           org_hyubjo_amt = org_hyubjo_amt + de_hyubjo_amt
           org_school_amt = org_school_amt + de_school_amt
           org_nhis_bla_amt = org_nhis_bla_amt + de_nhis_bla_amt
           org_long_bla_amt = org_long_bla_amt + de_long_bla_amt
		   org_deduct_tot = org_deduct_tot + de_deduct_tot
		   
		   team_nps_amt = team_nps_amt + de_nps_amt
           team_nhis_amt = team_nhis_amt + de_nhis_amt
           team_epi_amt = team_epi_amt + de_epi_amt
		   team_longcare_amt = team_longcare_amt + de_longcare_amt
           team_income_tax = team_income_tax + de_income_tax
           team_wetax = team_wetax + de_wetax
		   team_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
           team_year_wetax = sum_year_wetax + de_year_wetax
		   team_year_incom_tax2 = sum_year_incom_tax2 + de_year_incom_tax2
           team_year_wetax2 = sum_year_wetax2 + de_year_wetax2
           team_other_amt1 = team_other_amt1 + de_other_amt1
           team_sawo_amt = team_sawo_amt + de_sawo_amt
           team_hyubjo_amt = team_hyubjo_amt + de_hyubjo_amt
           team_school_amt = team_school_amt + de_school_amt
           team_nhis_bla_amt = team_nhis_bla_amt + de_nhis_bla_amt
           team_long_bla_amt = team_long_bla_amt + de_long_bla_amt
		   team_deduct_tot = team_deduct_tot + de_deduct_tot
							  
    %>    
    
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nps_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_epi_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_longcare_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_income_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_wetax,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_incom_tax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_year_wetax2,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_other_amt1,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_sawo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_school_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_nhis_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_long_bla_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_hyubjo_amt,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(de_deduct_tot,0)%></div></td>
    <td width="100"><div align="right" class="style1"><%=formatnumber(pmg_curr_pay,0)%></div></td>
  </tr>
	<%
	    Rs.MoveNext()
	loop
	
	sum_curr_pay = sum_give_tot - sum_deduct_tot
	team_curr_pay = team_give_tot - team_deduct_tot
	org_curr_pay = org_give_tot - org_deduct_tot
	
	%>
                 <tr>
                    <td colspan="8" bgcolor="#EEFFFF" align="center"><%=bi_team%>&nbsp;&nbsp;&nbsp;����</div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_pay_count,0)%>&nbsp;��</td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_base_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_meals_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_postage_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_re_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_overtime_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_car_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_position_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_custom_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_job_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_job_support,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_jisa_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_long_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_disabled_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_give_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_nps_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_nhis_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_epi_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_longcare_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_income_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_incom_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_incom_tax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_year_wetax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_other_amt1,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_sawo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_school_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_nhis_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_long_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_hyubjo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_deduct_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(team_curr_pay,0)%></div></td>
                 </tr>

                 <tr>
				    <td colspan="8" bgcolor="#EEFFFF" align="center"><%=bi_org%>&nbsp;&nbsp;&nbsp;����ΰ�</div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_pay_count,0)%>&nbsp;��</td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_base_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_meals_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_postage_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_re_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_overtime_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_car_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_position_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_custom_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_job_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_job_support,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_jisa_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_long_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_disabled_pay,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_give_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_nps_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_nhis_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_epi_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_longcare_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_income_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_incom_tax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_wetax,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_incom_tax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_year_wetax2,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_other_amt1,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_sawo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_school_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_nhis_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_long_bla_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_hyubjo_amt,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_deduct_tot,0)%></div></td>
                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(org_curr_pay,0)%></div></td>
                 </tr>
    
  <tr>    
    <th colspan="8" style=" border-top:1px solid #e3e3e3;"><div align="center" class="style1">�Ѱ�</div></th>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(pay_count,0)%>&nbsp;��</td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_base_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_meals_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_postage_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_re_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_overtime_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_car_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_position_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_custom_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_job_support,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_jisa_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_long_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_disabled_pay,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_give_tot,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_nps_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_nhis_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_epi_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_longcare_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_income_tax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_wetax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_year_wetax,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_year_incom_tax2,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_year_wetax2,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_other_amt1,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_sawo_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_school_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_nhis_bla_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_long_bla_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_hyubjo_amt,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_deduct_tot,0)%></div></td>
    <td width="100" style=" border-top:1px solid #e3e3e3;"><div align="right" class="style1"><%=formatnumber(sum_curr_pay,0)%></div></td>
  </tr>
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
