<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab(6)
dim pay_count(6,3)
dim overtime_pay(6,3)
dim give_amt(6,3)
dim re_pay(6,3)
dim give_tot(6,3)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))
to_yyyy = mid(cstr(to_date),1,4)
to_mm = mid(cstr(to_date),6,2)
to_dd = mid(cstr(to_date),9,2)

give_date = to_date '������

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
main_title = cstr(curr_yyyy) + "�� " + cstr(curr_mm) + "�� " + " �޿� �����񱳺м�"

  for i = 1 to 6
     com_tab(i) = ""
	 for j = 1 to 3
	    pay_count(i,j) = 0
		overtime_pay(i,j) = 0
		give_amt(i,j) = 0
		re_pay(i,j) = 0
		give_tot(i,j) = 0
     next
  next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

'����޿� ����
if view_condi = "��ü" then
          com_tab(1) = "���̿��������"
		  com_tab(2) = "�޵�"
		  com_tab(3) = "���̳�Ʈ����"
		  com_tab(4) = "����������ġ"
		  com_tab(5) = "�ڸ��Ƶ𿣾�"
		  com_tab(6) = "�հ�"
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1')"
	else	  
		  com_tab(1) = view_condi
		  com_tab(6) = "�հ�"
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
      for i = 1 to 6
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,1) = pay_count(i,1) + 1
				 pay_count(6,1) = pay_count(6,1) + 1
		         overtime_pay(i,1) = overtime_pay(i,1) + int(rs("pmg_overtime_pay"))
				 overtime_pay(6,1) = overtime_pay(6,1) + int(rs("pmg_overtime_pay"))
		         give_amt(i,1) = give_amt(i,1) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(6,1) = give_amt(6,1) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,1) = re_pay(i,1) + int(rs("pmg_re_pay"))
				 re_pay(6,1) = re_pay(6,1) + int(rs("pmg_re_pay"))
		         give_tot(i,1) = give_tot(i,1) + int(rs("pmg_give_total"))
				 give_tot(6,1) = give_tot(6,1) + int(rs("pmg_give_total"))
		  end if		 
	  next			 
	rs.movenext()
loop
rs.close()		

'���� �޿�
bef_month = mid(cstr(pmg_yymm),1,4) + mid(cstr(pmg_yymm),5,2)
bef_month = cstr(int(bef_month) - 1)
if mid(bef_month,5) = "00" then
	bef_year = cstr(int(mid(bef_month,1,4)) - 1)
	bef_month = bef_year + "12"
end if	

if view_condi = "��ü" then
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_month+"' ) and (pmg_id = '1')"
	else	  
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_month+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
      for i = 1 to 6
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,2) = pay_count(i,2) + 1
				 pay_count(6,2) = pay_count(6,2) + 1
		         overtime_pay(i,2) = overtime_pay(i,2) + int(rs("pmg_overtime_pay"))
				 overtime_pay(6,2) = overtime_pay(6,2) + int(rs("pmg_overtime_pay"))
		         give_amt(i,2) = give_amt(i,2) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(6,2) = give_amt(6,2) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,2) = re_pay(i,2) + int(rs("pmg_re_pay"))
				 re_pay(6,2) = re_pay(6,2) + int(rs("pmg_re_pay"))
		         give_tot(i,2) = give_tot(i,2) + int(rs("pmg_give_total"))
				 give_tot(6,2) = give_tot(6,2) + int(rs("pmg_give_total"))
		  end if		 
	  next			 
	rs.movenext()
loop
rs.close()		

'���� �޿�
bef_yearmon = cstr(int(mid(cstr(pmg_yymm),1,4)) - 1) + mid(cstr(pmg_yymm),5,2)
if view_condi = "��ü" then
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_yearmon+"' ) and (pmg_id = '1')"
	else	  
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_yearmon+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
      for i = 1 to 6
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,3) = pay_count(i,3) + 1
				 pay_count(6,3) = pay_count(6,3) + 1
		         overtime_pay(i,3) = overtime_pay(i,3) + int(rs("pmg_overtime_pay"))
				 overtime_pay(6,3) = overtime_pay(6,3) + int(rs("pmg_overtime_pay"))
		         give_amt(i,3) = give_amt(i,3) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(6,3) = give_amt(6,3) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,3) = re_pay(i,3) + int(rs("pmg_re_pay"))
				 re_pay(6,3) = re_pay(6,3) + int(rs("pmg_re_pay"))
		         give_tot(i,3) = give_tot(i,3) + int(rs("pmg_give_total"))
				 give_tot(6,3) = give_tot(6,3) + int(rs("pmg_give_total"))
		  end if		 
	  next			 
	rs.movenext()
loop
rs.close()		

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
<title>�� �޿����޴���</title>
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
     <td colspan="3" align="center" class="style24BC"><%=main_title%></td>
  </tr>
  <tr>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
  </tr>
  <tr>
	 <td width="33%" height="30" align="left"><span class="style14BC">�ۼ���&nbsp;:&nbsp; <%=curr_date%></span></td>
	 <td width="*" height="30" align="center"><span class="style14BC">&nbsp;</span></td>
	 <td width="33%" height="30" align="right"><span class="style14BC">����:��(����)</span></td>
  </tr>  
</table>
<table width="1030" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td colspan="2" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12BC">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</strong></td>
    <td height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12BC"><%=mid(pmg_yymm,1,4)%>��&nbsp;<%=mid(pmg_yymm,5,2)%>��</strong></td>
    <td height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12BC"><%=mid(bef_month,1,4)%>��&nbsp;<%=mid(bef_month,5,2)%>��</strong></td>
    <td height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12BC"><%=mid(bef_yearmon,1,4)%>��&nbsp;<%=mid(bef_yearmon,5,2)%>��</strong></td>
    <td height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#f8f8f8;"><strong class="style12BC">���</strong></td>
  </tr>
<%
    b_pay_count = 0
    b_overtime_pay = 0
    b_give_amt = 0
    b_re_pay = 0
    b_give_tot = 0
						
	y_pay_count = 0
    y_overtime_pay = 0
    y_give_amt = 0
    y_re_pay = 0
    y_give_tot = 0
						
  for i = 1 to 6 
   	if	com_tab(i) <> "" then
%>	  
  <tr>
    <td rowspan="5" width="*" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=com_tab(i)%></span></td>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ο�</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(pay_count(i,1),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(pay_count(i,2),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(pay_count(i,3),0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��Ư��</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(overtime_pay(i,1),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(overtime_pay(i,2),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(overtime_pay(i,3),0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�޿�</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(give_amt(i,1),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(give_amt(i,2),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(give_amt(i,3),0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ұ�</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(re_pay(i,1),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(re_pay(i,2),0)%>&nbsp;</span></td>
    <td width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(re_pay(i,3),0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <th width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�հ�</span></th>
    <th width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(give_tot(i,1),0)%>&nbsp;</span></th>
    <th width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(give_tot(i,2),0)%>&nbsp;</span></th>
    <th width="15%" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(give_tot(i,3),0)%>&nbsp;</span></th>
    <th width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></th>
  </tr>
<%
	end if
next
    b_pay_count = pay_count(6,1) - pay_count(6,2)
    b_overtime_pay = overtime_pay(6,1) - overtime_pay(6,2)
    b_give_amt = give_amt(6,1) - give_amt(6,2)
    b_re_pay = re_pay(6,1) - re_pay(6,2)
    b_give_tot = give_tot(6,1) - give_tot(6,2)
						
	y_pay_count = pay_count(6,1) - pay_count(6,3)
    y_overtime_pay = overtime_pay(6,1) - overtime_pay(6,3)
    y_give_amt = give_amt(6,1) - give_amt(6,3)
    y_re_pay = re_pay(6,1) - re_pay(6,3)
    y_give_tot = give_tot(6,1) - give_tot(6,3)
%>    
  <tr>
    <td rowspan="5" width="*" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�����������</span></td>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ο�</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(b_pay_count,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��Ư��</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(b_overtime_pay,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�޿�</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(b_give_amt,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ұ�</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(b_re_pay,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <th width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">������</span></th>
    <th colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(b_give_tot,0)%>&nbsp;</span></th>
    <th width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></th>
  </tr>
  <tr>
    <td rowspan="5" width="*" height="30" align="center"><span class="style12C">����������</span></td>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ο�</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(y_pay_count,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">��Ư��</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(y_overtime_pay,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�޿�</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(y_give_amt,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">�ұ�</span></td>
    <td colspan="3" height="30" align="right" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C"><%=formatnumber(y_re_pay,0)%>&nbsp;</span></td>
    <td width="20%" height="30" align="center" style=" border-bottom:1px solid #e3e3e3;"><span class="style12C">&nbsp;</span></td>
  </tr>
  <tr>
    <th width="15%" height="30" align="center"><span class="style12C">������</span></th>
    <th colspan="3" height="30" align="right"><span class="style12C"><%=formatnumber(y_give_tot,0)%>&nbsp;</span></th>
    <th width="20%" height="30" align="center"><span class="style12C">&nbsp;</span></th>
  </tr>
</table>
</p>	

</body>
</html>
