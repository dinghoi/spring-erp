<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sch_tab(10,10)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
in_tax_id = request("in_tax_id") 

curr_date = datevalue(mid(cstr(now()),1,10))
to_yyyy = mid(cstr(to_date),1,4)
to_mm = mid(cstr(to_date),6,2)
to_dd = mid(cstr(to_date),9,2)

give_date = to_date '������

tax_man_name = ""

if view_condi = "���̿��������" then
      company_name = "(��)" + "���̿��������"
	  owner_name = "�����"
	  addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif view_condi = "�޵�" then
              company_name = "(��)" + "�޵�"
			  owner_name = "������"
	          addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif view_condi = "���̳�Ʈ����" then
                     company_name = "���̳�Ʈ����" + "(��)"
					 owner_name = "���߿�"
	                 addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif view_condi = "����������ġ" then
                        company_name = "(��)" + "����������ġ"	
						owner_name = "�ڹ̾�"
	                    addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if 

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
main_title = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " �޿�����"

	sum_give_tot = 0
    sum_income_tax = 0
    sum_wetax = 0
	sum_year_incom_tax = 0
    sum_year_wetax = 0
	sum_special_tax = 0
	sum_deduct_tot = 0
	pay_count = 0	
	sum_curr_pay = 0
	
	a02_give_tot = 0
    a02_income_tax = 0
    a02_wetax = 0
	a02_count = 0	
	
	a03_give_tot = 0
    a03_income_tax = 0
    a03_wetax = 0
	a03_count = 0	
	
	a04_give_tot = 0
    a04_income_tax = 0
    a04_wetax = 0
	a04_count = 0	
	
	a10_give_tot = 0
    a10_income_tax = 0
    a10_wetax = 0
	a10_count = 0	
	
	a21_give_tot = 0
    a21_income_tax = 0
    a21_wetax = 0
	a21_count = 0	
	
	a22_give_tot = 0
    a22_income_tax = 0
    a22_wetax = 0
	a22_count = 0	
	
	a20_give_tot = 0
    a20_income_tax = 0
    a20_wetax = 0
	a20_count = 0	
	
	sum_alba_give_total = 0
    sum_tax_amt1 = 0
    sum_tax_amt2 = 0
	sum_deduct_tot = 0
	
	a32_give_tot = 0
    a32_income_tax = 0
    a32_wetax = 0
	a32_count = 0	
	
	a30_give_tot = 0
    a30_income_tax = 0
    a30_wetax = 0
	a30_count = 0
	
	tot_give_tot = 0
    tot_income_tax = 0
    tot_wetax = 0
	tot_year_incom_tax = 0
    tot_year_wetax = 0
	tot_special_tax = 0
	tot_deduct_tot = 0
	tot_pay_count = 0	
	tot_curr_pay = 0		

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

'�ٷμҵ�
Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    emp_no = rs("pmg_emp_no")
    pmg_give_tot = rs("pmg_give_total")
    pay_count = pay_count + 1
				  
    sub_give_hap = int(rs("pmg_postage_pay")) + int(rs("pmg_re_pay")) + int(rs("pmg_car_pay")) + int(rs("pmg_position_pay")) + int(rs("pmg_custom_pay")) + int(rs("pmg_job_pay")) + int(rs("pmg_job_support")) + int(rs("pmg_jisa_pay")) + int(rs("pmg_long_pay")) + int(rs("pmg_disabled_pay"))
	
	sum_give_tot = sum_give_tot + int(rs("pmg_give_total"))

    Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
    if not Rs_dct.eof then

            de_income_tax = int(Rs_dct("de_income_tax"))
            de_wetax = int(Rs_dct("de_wetax"))
			de_year_incom_tax = int(Rs_dct("de_year_incom_tax"))
            de_year_wetax = int(Rs_dct("de_year_wetax"))
		    de_deduct_tot = int(Rs_dct("de_deduct_total"))	
	     else
            de_income_tax = 0
            de_wetax = 0
			de_year_incom_tax = 0
            de_year_wetax = 0
		    de_deduct_tot = 0
     end if
     Rs_dct.close()
	 
     sum_income_tax = sum_income_tax + de_income_tax
     sum_wetax = sum_wetax + de_wetax
	 sum_year_incom_tax = sum_year_incom_tax + de_year_incom_tax
     sum_year_wetax = sum_year_wetax + de_year_wetax
	 sum_deduct_tot = sum_deduct_tot + de_deduct_tot

	rs.movenext()
loop
rs.close()

a10_give_tot = sum_give_tot + a02_give_tot + a03_give_tot + a03_give_tot 
a10_income_tax = sum_income_tax + a02_income_tax + a03_income_tax + a04_income_tax
a10_wetax = sum_wetax + a02_wetax + a03_wetax + a04_wetax
a10_count = pay_count + a02_count + a03_count + a04_count

'�����ҵ�
a20_give_tot = a21_give_tot + a22_give_tot
a20_income_tax = a21_income_tax + a22_income_tax
a20_wetax = a21_wetax + a22_wetax
a20_count = a21_count + a22_count

'����ҵ�
Sql = "select * from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    alba_count = alba_count + 1
				  
    sum_alba_give_total = sum_alba_give_total + int(rs("alba_give_total"))
    sum_tax_amt1 = sum_tax_amt1 + int(rs("tax_amt1"))
    sum_tax_amt2 = sum_tax_amt2 + int(rs("tax_amt2"))
	sum_deduct_tot = sum_deduct_tot + (int(rs("tax_amt1")) + int(rs("tax_amt2")) + int(rs("de_other")))
	
	rs.movenext()
loop
rs.close()

a30_give_tot = sum_alba_give_total + a32_give_tot
a30_income_tax = sum_tax_amt1 + a32_income_tax
a30_wetax = sum_tax_amt2 + a32_wetax
a30_count = alba_count + a32_count

'�Ѱ�
tot_give_tot = a10_give_tot + a20_give_tot + a30_give_tot
tot_income_tax = a10_income_tax + a20_income_tax + a30_income_tax
tot_wetax = a10_wetax + a20_wetax + a30_wetax
tot_pay_count = a10_count + a20_count + a30_count

if in_tax_id = "1" then 
   tax_id_name = "����Ű�" 
   elseif in_tax_id = "2" then 
          tax_id_name = "�б�" 
          elseif in_tax_id = "3" then 
		         tax_id_name = "����" 
end if

title_line = " ��ҵ漼�� �����Ģ[���� ��21ȣ����]<����2014.3.14> "

Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') ORDER BY pmg_company,pmg_org_code,pmg_emp_no ASC"
Rs.Open Sql, Dbconn, 1

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>�޿����� �ý���</title>
        <script src="/java/common.js" type="text/javascript"></script>
        <script type="text/javascript">
	function printWindow(){
//		viewOff("button");   
		factory.printing.header = ""; //�Ӹ��� ����
		factory.printing.footer = ""; //������ ����
		factory.printing.portrait = true; //��¹��� ����: true - ����, false - ����
		factory.printing.leftMargin = 13; //���� ���� ����
		factory.printing.topMargin = 25; //���� ���� ����
		factory.printing.rightMargin = 13; //�����P ���� ����
		factory.printing.bottomMargin = 15; //�ٴ� ���� ����
//		factory.printing.SetMarginMeasure(2); //�׵θ� ���� ������ ������ ��ġ�� ����
//		factory.printing.printer = ""; //������ �� ������ �̸�
//		factory.printing.paperSize = "A4"; //��������
//		factory.printing.pageSource = "Manusal feed"; //���� �ǵ� ���
//		factory.printing.collate = true; //������� ����ϱ�
//		factory.printing.copies = "1"; //�μ��� �ż�
//		factory.printing.SetPageRange(true,1,1); //true�� �����ϰ� 1,3�̸� 1���� 3������ ���
//		factory.printing.Printer(true); //����ϱ�
		factory.printing.Preview(); //�����츦 ���ؼ� ���
		factory.printing.Print(false); //�����츦 ���ؼ� ���
	}
	function printW() {
        window.print();
    }
	function goBefore () {
		history.back() ;
	}
	
</script>
<title>��õ¡�������Ȳ�Ű�</title>
<style type="text/css">
<!--
    	.style10C {font-size: 10px; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style10BC {font-size: 10px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style14BC {font-size: 14px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style24BC {font-size: 24px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
		.style10L {font-size: 8px; font-family: "����ü", "����ü", Seoul; text-align: left; }
-->
</style>
<style media="print"> 
.noprint     { display: none }
</style>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<div class="noprint">
<p><a href="#" onClick="printWindow()"><img src="image/printer.jpg" width="39" height="36" border="0" alt="����ϱ�" /></a></p>
</div>
<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>    
   
<table width="1030" cellpadding="0" cellspacing="0">
  <tr>
    <td width="33%" height="20" align="left"><span class="style12L"><%=title_line%></span></td>
    <td width="*" height="20" align="center"><span class="style12L">&nbsp;&nbsp;</span></td>
    <td width="33%" height="20" align="right"><span class="style12L">[���ڽŰ������]</span></td>
  </tr>  
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="6" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �Ű���</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:0px solid #ffffff;"><strong class="style14BC">��&nbsp;&nbsp; ��õ¡�������Ȳ�Ű�<br>��&nbsp;&nbsp; ��õ¡������ȯ�޽�û��</strong></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �ͼӳ��</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=curr_yyyy%>��<%=curr_mm%>��</span></td>
  </tr>
  <tr>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ſ�</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ݱ�</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ҵ�ó��</span></td>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">ȯ�޽�û</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� ���޿���</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=curr_yyyy%>��<%=curr_mm%>��</span></td>
  </tr>
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td rowspan="4" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��õ<br>¡��<br>�ǹ���</span></td>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���θ�(��ȣ)</span></td>
    <td rowspan="2" width="20%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=company_name%></span></td>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��ǥ��<br>(����)</span></td>
    <td rowspan="2" width="20%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=owner_name%></span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ϰ����� ����</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� , ��</span></td>
  </tr>
  <tr>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����ڴ�����������</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� , ��</span></td>
  </tr>
  <tr>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�����(�ֹ�)<br>��Ϲ�ȣ</span></td>
    <td rowspan="2" width="20%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=trade_no%></span></td>
    <td rowspan="2" width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�����<br>������</span></td>
    <td rowspan="2" width="20%" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=addr_name%></span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��ȭ��ȣ</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=tel_no%></span></td>
  </tr>
  <tr>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���ڿ����ּ�</span></td>
    <td width="15%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=e_mail%></span></td>
  </tr>
  <tr>
    <td colspan="7" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><strong class="style12C">1. ��õ¡���� �� ���μ��� (����: ��)</strong></td>
  </tr>
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td rowspan="3" colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����ҵ��ڱ���</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ڵ�</span></td>
    <td colspan="5" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��õ¡����</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �������<br>ȯ�޼���</span></td>
    <td colspan="5" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���μ���</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ҵ�����<br>(�����̴�,�Ϻκ��������)</span></td>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">¡������</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �ҵ漼��<br>(���꼼����)</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �����<br>Ư����</span></td>
  </tr>
  <tr>
    <td width="6%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �ο�</span></td>
    <td width="12%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �����޾�</span></td>
    <td width="12%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �ҵ漼��</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �����Ư����</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� ���꼼</span></td>
  </tr>
  <tr>
    <td rowspan="22" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��<br>��</span></td>
    <td rowspan="5" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��<br>
    ��<br>��<br>��</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���̼���</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A01</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(pay_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ߵ����</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A02</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a02_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a02_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a02_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�Ͽ�ٷ�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A03</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A04</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a04_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a03_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A10</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a10_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="3" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��<br>��<br>��<br>��</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���ݰ���</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A21</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a21_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a21_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a21_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�׿�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A22</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a22_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A20</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a20_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="3" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��<br>��<br>��<br>��</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ſ�¡��</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A25</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(alba_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_alba_give_total,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_tax_amt1,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A26</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a32_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a32_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a32_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A30</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(a30_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="3" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��<br>Ÿ<br>��<br>��</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���ݰ���</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A41</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�׿�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A42</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A40</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td rowspan="4" width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��<br>��<br>��<br>��</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���ݰ���</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A48</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��������(�ſ�)</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A45</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A46</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">������</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A47</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���ڼҵ�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A50</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���ҵ�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A60</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style10C">����������¡���׵�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A69</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">������ھ絵�ҵ�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A70</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="3%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���ܱ����ο�õ</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A80</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�����Ű�(����)</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A90</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><span class="style12C">&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���հ�</span></td>
    <td width="4%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">A99</span></td>
    <td width="6%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_pay_count,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_give_tot,0)%>&nbsp;</span></td>
    <td width="12%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(tot_income_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="12" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><strong class="style12C">2. ȯ�޼��� ���� (����: ��)</strong></td>
  </tr>
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���� ��ȯ�� ������ ���</span></td>
    <td colspan="4" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��� �߻� ȯ�޼���</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">?�������<br>ȯ�޼���<br>(��+��+<br>?+?)</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? �������<br>ȯ�޼���</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? �����̿�<br>ȯ�޼���<br>(?-?)</span></td>
    <td rowspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? ȯ�޽�û��</span></td>
  </tr>
  <tr>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� ����<br>��ȯ�޼���</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� ��ȯ��<br>��û����</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �����ܾ�<br>(��-��)</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�� �Ϲ�ȯ��</span></td>
    <td rowspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? ��Ź���<br>(����ȸ���)</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">? �׹���ȯ�޼���</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����ȸ���</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�պ� ��</span></td>
  </tr>
  <tr>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="10%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
    <td width="9%" height="20" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(sum_special_tax,0)%>&nbsp;</span></td>
  </tr>
</table>

<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td rowspan="11" width="70%" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;&nbsp;&nbsp;&nbsp;��õ¡���ǹ��ڴ� ���ҵ漼�� ����ɡ� ��185����1�׿����� ���� ������ �����ϸ�, <strong><br>&nbsp;&nbsp;�� ������ ����� �����Ͽ��� ��õ¡���ǹ��ڰ� �˰� �ִ� ��� �״�<br>&nbsp;&nbsp;�θ� ��Ȯ�ϰ� �������� Ȯ���մϴ�</strong>
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <%=mid(cstr(now()),1,4)%>��&nbsp;<%=mid(cstr(now()),6,2)%>��&nbsp;<%=mid(cstr(now()),9,2)%>��&nbsp;&nbsp;&nbsp;
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    ��õ¡���ǹ���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=company_name%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(���� �Ǵ� ��)&nbsp;&nbsp;&nbsp;
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=owner_name%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
    <br>
    <br>
    &nbsp;&nbsp;&nbsp;&nbsp;<strong>�����븮���� ���������ڰ��ڷμ� �� �Ű��� �����ϰ� �����ϰ�<br>&nbsp;&nbsp;�ۼ��Ͽ����� Ȯ���մϴ�</strong>
    <br>
    <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    �����븮��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=tax_man_name%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(���� �Ǵ� ��)&nbsp;&nbsp;&nbsp;
    <br>
    <br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��õ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>��������</strong>&nbsp;&nbsp;����
    </span>
    </td>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�Ű� ������ �ۼ�����<br>�� �ش���� "0"ǥ�ø� �մϴ�</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��ǥ(4-5)��</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">ȯ��(7�ʡ�9��)</span></td>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�°��(10��)</span></td>
  </tr>
  <tr>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�����븮��</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����ڵ�Ϲ�ȣ</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��ȭ��ȣ</span></td>
    <td colspan="2" height="20" align="left" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</span></td>
  </tr>
  <tr>
    <td colspan="3" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����ȯ�ޱݰ��½Ű�<br>��ȯ�ޱݾ� 2õ���� �̸��� ��쿡�� �����ϴ�</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">����ó</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��������</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">���¹�ȣ</span></td>
    <td colspan="2" height="20" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
</table>
<table width="1030" cellpadding="0" cellspacing="0">
  <tr>
    <td width="33%" height="20" align="left"><span class="style10L">&nbsp;&nbsp;</span></td>
    <td width="*" height="20" align="center"><span class="style10L">&nbsp;&nbsp;</span></td>
    <td width="33%" height="20" align="right"><span class="style12L">210����297��(�����80g/��)</span></td>
  </tr>
</table>  
</p>	
</body>
</html>
