<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

from_date = Request("from_date")
to_date = Request("to_date")
view_condi = request("view_condi")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

main_title = from_date + " �� " + to_date + " �Ի��� ��Ȳ"

if view_condi = "��ü" then
   Sql = "select * from emp_master where emp_in_date >= '"+from_date+"' and emp_in_date <= '"+to_date+"' ORDER BY emp_no,emp_name ASC"
   else  
   Sql = "select * from emp_master where emp_company = '"+view_condi+"' and emp_in_date >= '"+from_date+"' and emp_in_date <= '"+to_date+"' ORDER BY emp_no,emp_name ASC"
end if
Rs.Open Sql, Dbconn, 1


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript">
            function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //�Ӹ��� ����
                factory.printing.footer = ""; //������ ����
                factory.printing.portrait = false; //��¹��� ����: true - ����, false - ����
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
        </script>
		<style type="text/css">
        <!--
    	    .style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
    	    .style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
            .style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
            .style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }            
			.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
            .style16BC {font-size: 16px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
            .style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
            .style14C {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: center; }
            .style14BC {font-size: 14px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
            .style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
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
	    <td width="5%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">���</span></td>
	    <td width="5%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">��  ��</span></td>
	    <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�������</span></td>
	    <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">����</span></td>
	    <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">����</span></td>
	    <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">��å</span></td>
        <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�Ի���</span></td>
        <td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�Ҽ�</span></td>
        <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�����з�</span></td>
        <td width="8%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">��ֿ���</span></td>
        <td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">����óȸ��</span></td>
        <td width="28%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">����</span></td>
      </tr>
	<% 
		do until rs.eof
     
	 	    if rs("emp_org_baldate") = "1900-01-01" then
			   emp_org_baldate = ""
			   else 
			   emp_org_baldate = rs("emp_org_baldate")
			end if
			
			if rs("emp_grade_date") = "1900-01-01" then
			   emp_grade_date = ""
			   else
			   emp_grade_date = rs("emp_grade_date")
			end if
			
	%>					
	  <tr>
	    <td width="5%" height="30" align="center"><span class="style12C"><%=rs("emp_no")%></span></td>
	    <td width="5%" height="30" align="center"><span class="style12C"><%=rs("emp_name")%></span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("emp_birthday")%></span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("emp_grade")%></span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("emp_job")%></span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("emp_position")%>&nbsp;</span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("emp_in_date")%></span></td>
	    <td width="9%" height="30" align="center"><span class="style12C"><%=rs("emp_org_name")%></span></td>
        <td width="6%" height="30" align="center"><span class="style12C"><%=rs("emp_last_edu")%>&nbsp;</span></td>
        <td width="8%" height="30" align="center"><span class="style12C"><%=rs("emp_disabled")%>&nbsp;<%=rs("emp_disab_grade")%></span></td>
        <td width="9%" height="30" align="center"><span class="style12C"><%=rs("emp_reside_company")%>&nbsp;</span></td>
        <td width="28%" height="30" align="left"><span class="style12C"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></span></td>
      </tr>
	<% 
		    rs.movenext()
	    loop
		rs.close()
	%>
    </table>
</body>
</html>

