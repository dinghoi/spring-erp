<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

from_date = Request("from_date")
to_date = Request("to_date")
view_condi = request("view_condi")
app_id=Request("app_id")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

main_title = view_condi + "(" + app_id + ") - �λ�߷� ��Ȳ(" + from_date + " �� " + to_date + ")"

if view_condi = "��ü" then
   if app_id = "��ü" then
           Sql = "select * from emp_appoint where app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	  else
		   Sql = "select * from emp_appoint where app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
   end if
   else
      if app_id = "��ü" then
	          Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
		 else
			  Sql = "select * from emp_appoint where app_to_company = '"+view_condi+"' and app_id = '"+app_id+"' and app_date >= '"+from_date+"' and app_date <= '"+to_date+"'  and (app_empno < '900000') ORDER BY app_date,app_empno ASC"
	  end if
end if
Rs.Open Sql, Dbconn, 1
%>
<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">-->
<!DOCTYPE HTML>
<html lang="ko">
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
    <p><a href="#" onClick="printWindow()"><img src="/image/printer.jpg" width="39" height="36" border="0" alt="����ϱ�" /></a></p>
    </div>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>
<table width="1250" border="0" cellspacing="0" cellpadding="0">
<tr>
      <td colspan="3" align="center" class="style32BC"><%=main_title%></td>
      </tr>
	  <tr>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
      </tr>
</table>
<table width="1250" border="1" cellspacing="0" cellpadding="0">
	  <tr>
         <td colspan="5" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; bgcolor=#BFBFFF"><strong class="style12BC">�߷ɻ���</strong></td>
         <td colspan="3" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;"><strong class="style12BC">�߷���</strong></td>
         <td colspan="4" height="30" align="center" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;"><strong class="style12BC">�߷���</strong></td>
      </tr>
      <tr>
	    <td width="5%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">���</span></td>
	    <td width="5%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">����</span></td>
	    <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�߷���</span></td>
	    <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�߷ɱ���</span></td>
	    <td width="6%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�߷�����</span></td>
	    <td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">ȸ��</span></td>
        <td width="10%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�Ҽ�</span></td>
        <td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">����/å</span></td>
        <td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">ȸ��</span></td>
        <td width="10%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�Ҽ�</span></td>
        <td width="9%" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">����/å</span></td>
        <td width="*" height="30" align="center" bgcolor="#BFBFFF"><span class="style12C">�߷ɳ���</span></td>
      </tr>
	<%
		do until rs.eof

	%>
	  <tr>
	    <td width="5%" height="30" align="center"><span class="style12C"><%=rs("app_empno")%></span></td>
	    <td width="5%" height="30" align="center"><span class="style12C"><%=rs("app_emp_name")%></span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("app_date")%></span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("app_id")%></span></td>
	    <td width="6%" height="30" align="center"><span class="style12C"><%=rs("app_id_type")%>&nbsp;</span></td>
	    <td width="9%" height="30" align="center"><span class="style12C"><%=rs("app_to_company")%>&nbsp;</span></td>
	    <td width="10%" height="30" align="center"><span class="style12C"><%=rs("app_to_org")%>(<%=rs("app_to_orgcode")%>)</span></td>
	    <td width="9%" height="30" align="center"><span class="style12C"><%=rs("app_to_grade")%>-<%=rs("app_to_position")%></span></td>
        <td width="9%" height="30" align="center"><span class="style12C"><%=rs("app_be_company")%>&nbsp;</span></td>
        <td width="10%" height="30" align="center"><span class="style12C"><%=rs("app_be_org")%>(<%=rs("app_be_orgcode")%>)</span></td>
        <td width="9%" height="30" align="center"><span class="style12C"><%=rs("app_be_grade")%>-<%=rs("app_be_position")%></span></td>
        <td width="*" height="30" align="left"><span class="style12C"><%=rs("app_start_date")%>&nbsp;-&nbsp;<%=rs("app_finish_date")%>&nbsp;<%=rs("app_be_enddate")%>&nbsp;<%=rs("app_reward")%>&nbsp;:&nbsp;<%=rs("app_comment")%></span></td>
      </tr>
	<%
		    rs.movenext()
	    loop
		rs.close()
	%>
    </table>
</body>
</html>

