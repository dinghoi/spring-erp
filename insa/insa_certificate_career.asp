<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'on Error resume next

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
Dim emp_user, curr_date, curr_year, curr_month, curr_day
Dim emp_name, cfm_use, cfm_use_dept, cfm_comment
Dim rsCert, rsMax
Dim companyAddr

curr_date = Mid(CStr(Now()),1,10)
curr_year = Mid(CStr(Now()),1,4)
curr_month = Mid(CStr(Now()),6,2)
curr_day = Mid(CStr(Now()),9,2)

emp_no = Request.Form("in_empno")
emp_name = Request.Form("in_name")

cfm_use = Request.Form("cfm_use")
cfm_use_dept = Request.Form("cfm_use_dept")
cfm_comment = Request.Form("cfm_comment")

objBuilder.Append "SELECT emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
objBuilder.Append "	emtt.emp_org_name, emtt.emp_name, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_person1, emtt.emp_person2, emtt.emp_in_date, emtt.emp_sido, "
objBuilder.Append "	emtt.emp_gugun, emtt.emp_dong, emtt.emp_addr, emp_end_date, "
objbuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emp_no = '" & emp_no  & "' "

Set rsCert = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCert.eof Then
	'emp_company = rsCert("emp_company")
	emp_company = rsCert("org_company")
Else
	emp_company = ""
End If

Select Case emp_company
	Case "���̿�", "���̿��������" : emp_company = "(��)" & "���̿�"
	Case "���̳�Ʈ����" : emp_company = "(��)" & "���̳�Ʈ����"
	Case "���̽ý���", "�ڸ��Ƶ𿣾�" : emp_company = "(��)" & "���̽ý���"
	Case "����������ġ" : emp_company = "(��)" & "����������ġ"
	Case "�޵�" : emp_company = "(��)" & "�޵�"
End Select

Dim cfm_company, cfm_emp_name, cfm_org_name, cfm_job, cfm_position
Dim cfm_person1, cfm_person2, emp_in_date, emp_in_year
Dim emp_in_month, emp_in_day, year_cnt, mon_cnt, day_cnt, seq_last
Dim cfm_number, cfm_type, max_seq, cfm_seq, emp_person2, emp_end_date
Dim target_date

cfm_company = rsCert("org_company")

Select Case cfm_company
	Case "���̿��������" : cfm_company = "���̿�"
	Case "�ڸ��Ƶ𿣾�" : cfm_company = "���̽ý���"
End Select

cfm_emp_name = rsCert("emp_name")
cfm_org_name = rsCert("org_name")
cfm_job = rsCert("emp_job")
cfm_position = rsCert("emp_position")
cfm_person1 = rsCert("emp_person1")
cfm_person2 = rsCert("emp_person2")

emp_in_date = Mid(CStr(rsCert("emp_in_date")), 1, 10)
emp_in_year = Mid(CStr(rsCert("emp_in_date")), 1, 4)
emp_in_month = Mid(CStr(rsCert("emp_in_date")), 6, 2)
emp_in_day = Mid(CStr(rsCert("emp_in_date")), 9, 2)

If rsCert("emp_end_date") = "1900-01-01" Or IsNull(rsCert("emp_end_date")) Then
   emp_end_date = ""
   target_date = curr_date
Else
   emp_end_date = rsCert("emp_end_date")
   target_date = rsCert("emp_end_date")
End If

year_cnt = DateDiff("yyyy", rsCert("emp_in_date"), target_date)
mon_cnt = DateDiff("m", rsCert("emp_in_date"), target_date)
day_cnt = DateDiff("d", rsCert("emp_in_date"), target_date)

Dim y_cnt, m_cnt, stay_tit

year_cnt = Int(year_cnt) + 1
mon_cnt = Int(mon_cnt) + 1
day_cnt = Int(day_cnt) + 1
y_cnt = Int(mon_cnt / 12)
m_cnt = mon_cnt - (y_cnt * 12)

If y_cnt > 0 And m_cnt > 0 Then
	stay_tit = CStr(y_cnt) & "�� " & cstr(m_cnt) & "����"
ElseIf y_cnt > 0 And m_cnt = 0 Then
	stay_tit = CStr(y_cnt) & "�� "
ElseIf y_cnt = 0 and m_cnt > 0 Then
	stay_tit = CStr(m_cnt) & "���� "
ElseIf y_cnt = 0 and m_cnt = 0 Then
	stay_tit = CStr(m_cnt) & "���� "
End If

seq_last = ""
cfm_number = curr_year
cfm_type = "�������"

'sql="select max(cfm_seq) as max_seq from emp_confirm where cfm_type = '"&cfm_type&"' and cfm_number = '"&curr_year&"'"
objBuilder.Append "SELECT MAX(cfm_seq) AS max_seq "
objBuilder.Append "FROM emp_confirm "
objBuilder.Append "WHERE cfm_type = '"&cfm_type&"' "
objBuilder.Append "	AND cfm_number = '"&curr_year&"' "

Set rsMax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsMax("max_seq")) Then
	seq_last = "0001"
Else
	max_seq = "000" & CStr((int(rsMax("max_seq")) + 1))
	seq_last = right(max_seq,4)
End If
rsMax.close() : Set rsMax = Nothing

cfm_seq = seq_last
'emp_person2 = "*******"
emp_person2 = cfm_person2
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript">
	//ActiveX ������� IE11 �� ���������� ���� �߻�(scriptX �̻�� ó��)[����ȣ_20220204]
	function printWindow(){
//		viewOff("button");
		factory.printing.header = ""; //�Ӹ��� ����
		factory.printing.footer = ""; //������ ����
		factory.printing.portrait = true; //��¹��� ����: true - ����, false - ����
		factory.printing.leftMargin = 5; //���� ���� ����
		factory.printing.topMargin = 15; //���� ���� ����
		factory.printing.rightMargin = 5; //�����P ���� ����
		factory.printing.bottomMargin = 0; //�ٴ� ���� ����
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
	function printW(){
        window.print();
    }

	function goBefore(){
		//history.back() ;
		location.href = "/insa/insa_confirm_list.asp";
	}

	//����Ʈ �Լ� �ű� �ۼ�[����ȣ_20220204]
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
<title>������� ���</title>
	<style type="text/css">
    <!--
    	.style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style1 {font-size:14px;color: #666666}
		.style2 {font-size:12px;color: #666666}
    -->
    </style>
    <style media="print">
    .noprint     { display: none }
    </style>
    <style type="text/css" media="screen">
    .onlyprint {display:; }
    </style>

	</head>

    <body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
    <div align=center class="noprint">
     <p>
        <a href="javascript:fnPrint();"><img src="/image/b_print.gif" border="0" /></a>
        <a href="javascript:goBefore();"><img src="/image/b_close.gif" border="0" /></a>
     </p>
    </div>
    <!--<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
    </object>-->
	<div id="print_pg">
        <table width="750" border="1" cellspacing="10" cellpadding="0" align="center" class="onlyprint" style="border:10px solid #0072BE;">
          <tr>
             <td width="100%" height="100%" bgcolor="ffffff" align="center" valign="top" style="padding-left:20px; padding-right:20px;" >
	             <table width="100%" border="0" cellspacing="0" cellpadding="0">
	               <tr>
		             <td align="left" height="60" valign="middle" style="padding-left:20px;" >��<%=cfm_number%>��<%=cfm_seq%>&nbsp;ȣ</td>
	               </tr>
	               <tr>
		             <td height="130" align="center" valign="middle"><strong class="style32BC">�� �� �� �� ��</strong></td>
	               </tr>
	               <tr>
		             <td valign="middle" align="center">
		               <table width="600" cellspacing="1" cellpadding="12"  style="border:1px solid #000000;">
                         <tr>
                            <td width="22%" height="30" align="center" valign="middle" style="border-bottom:1px solid #000000; border-right:1px solid #000000; background-color:#eaeaea;"><span class="style1">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</span></td>
                            <td width="28%" align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=rsCert("emp_name")%></strong></td>
                            <td width="22%" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">�ֹε�Ϲ�ȣ</span></td>
                            <td width="28%" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=rsCert("emp_person1")%>-<%=emp_person2%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=emp_company%></strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� </span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=rsCert("emp_job")%></strong></td>
                         </tr>
                         <tr>
                         <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">���ٹ����Ի���</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=mid(cstr(rsCert("emp_in_date")),1,4)%>��&nbsp;<%=mid(cstr(rsCert("emp_in_date")),6,2)%>��&nbsp;<%=mid(cstr(rsCert("emp_in_date")),9,2)%>��</strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">��&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;�� </span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=mid(cstr(emp_end_date),1,4)%>��&nbsp;<%=mid(cstr(emp_end_date),6,2)%>��&nbsp;<%=mid(cstr(emp_end_date),9,2)%>��</strong></td>
                         </tr>
                         <tr>
                         <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;��</span></td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000;"><span class="style1"><strong><%=stay_tit%></strong></td>
                            <td align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</span>&nbsp;</td>
                            <td align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=cfm_use%>&nbsp;-<%=cfm_use_dept%></strong></td>
                         </tr>
                         <tr>
                            <td height="30" align="center" valign="middle" style="border-bottom:1px solid #000000;border-right:1px solid #000000; background-color:#EAEAEA; "><span class="style1">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</span></td>
                            <td colspan="3" align="left" valign="middle" style="border-bottom:1px solid #000000;"><span class="style1"><strong><%=rsCert("emp_sido")%>&nbsp;<%=rsCert("emp_gugun")%>&nbsp;<%=rsCert("emp_dong")%>&nbsp;<%=rsCert("emp_addr")%></strong></td>
                         </tr>
                        <tr>
                           <td height="30" align="center" valign="middle" style="border-right:1px solid #000000; background-color:#EAEAEA;"><span class="style1">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</span></td>
                           <td colspan="3"><span class="style1"><strong><%=cfm_comment%></strong></td>
                       </tr>
                </table></td>
	       </tr>
	       <tr>
		      <td height="280" align="center"><font style="font-size:16px"><strong>�� ������ ������� ������</td>
	       </tr>
	       <tr>
		   <%
			Select Case cfm_company
				Case "���̿�" : companyAddr = "����� ��õ�� ���������2�� 14, �븢��ũ��Ÿ��12�� 1405ȣ"
				Case "���̳�Ʈ����" : companyAddr = "����� ��õ�� ���������2�� 18, �븢��ũ��Ÿ��1�� 605ȣ"
				Case "���̽ý���" : companyAddr = "����� ��õ�� ���������2�� 18, �븢��ũ��Ÿ��1�� 406ȣ"
				Case Else
					companyAddr = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
			End Select
		  %>
          <%' if cfm_company = "���̿�" Or cfm_company = "���̳�Ʈ����" then %>
		      <td height="60" align="right" width="600"><font style="font-size:14px"><%=Mid(CStr(Now()), 1, 4)%>��&nbsp;<%=Mid(CStr(Now()), 6, 2)%>��&nbsp;<%=Mid(CStr(Now()), 9, 2)%>��<br/><br/>
			  <%=companyAddr%>
			  </td>
          <%' else %>
              <!--<td height="60" align="right" width="600"><font style="font-size:14px"><%=mid(cstr(now()),1,4)%>��&nbsp;<%=mid(cstr(now()),6,2)%>��&nbsp;<%=mid(cstr(now()),9,2)%>��<br/><br/>
				����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)</td>-->
          <%' end if %>
	      </tr>
	      <tr>
          <%
		  if cfm_company = "���̿�" then
		  %>
	         <!--<td height="60" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>�ֽ�ȸ�� ���̿��������<br />-->
			 <td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_one_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>�ֽ�ȸ�� ���̿�<br />
			<font style="font-size:14px">��ǥ�̻� </font><font style="font-size:16px"><b>�����</b></font></td>
          <% end if %>
          <% if cfm_company = "�޵�" then %>
	        <td height="60" align="right" valign="bottom" width="100%"><img src="/image/k_hudis001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>�ֽ�ȸ�� �޵�<br />
			<font style="font-size:14px">��ǥ�̻� </font><font style="font-size:16px"><b>�ڿ���</b></font></td>
          <% end if %>
          <% if cfm_company = "���̳�Ʈ����" then %>
	        <!--<td height="60" align="right" valign="middle" width="100%"><img src="image/k_net001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>���̳�Ʈ���� �ֽ�ȸ��<br />-->
			<td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_net_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>�ֽ�ȸ�� ���̳�Ʈ����<br />
			<!--<font style="font-size:14px">��ǥ�̻� </font><font style="font-size:16px"><b>�����</b></font><br/>-->
			<font style="font-size:14px">��ǥ�̻� </font><font style="font-size:16px"><b>�̵���</b></font></td>
          <% end if %>
          <% if cfm_company = "����������ġ" then %>
	        <td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_one_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>�ֽ�ȸ�� ����������ġ<br />
			<font style="font-size:14px">��ǥ�̻� </font><font style="font-size:16px"><b>�ڹ̾�</b></font></td>
          <% end if %>
          <% if cfm_company = "���̽ý���" then %>
	        <!--<td height="60" align="right" valign="middle" width="100%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right><font style="font-size:14px"><br><br>�ڸ��Ƶ𿣾� �ֽ�ȸ��<br />-->
			<td height="60" align="right" valign="bottom" width="100%"><img src="/image/stamp/k_sys_2021_001.png" width="80" height="80" alt="" align="right"><font style="font-size:14px"><br><br>�ֽ�ȸ�� ���̽ý���<br />
			<font style="font-size:14px">��ǥ�̻� </font><font style="font-size:16px"><b>�۰���</b></font></td>
          <% end if %>
	      </tr>
	   </table>
	<br><br><br>


	   </td>
    </tr>

 <%
 		'sql = "insert into emp_confirm(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type,cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position,cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment,cfm_reg_date,cfm_reg_user) values "
		'sql = sql +	" ('"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"','"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"','"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"',now(),'"&emp_user&"')"

		objBuilder.Append "INSERT INTO emp_confirm(cfm_empno,cfm_number,cfm_seq,cfm_date,cfm_type, "
		objBuilder.Append "	cfm_emp_name,cfm_company,cfm_org_name,cfm_job,cfm_position, "
		objBuilder.Append "	cfm_person1,cfm_person2,cfm_use,cfm_use_dept,cfm_comment, "
		objBuilder.Append "	cfm_reg_date,cfm_reg_user)VALUES("
		objBuilder.Append "'"&emp_no&"','"&cfm_number&"','"&cfm_seq&"','"&curr_date&"','"&cfm_type&"', "
		objBuilder.Append "'"&cfm_emp_name&"','"&cfm_company&"','"&cfm_org_name&"','"&cfm_job&"','"&cfm_position&"', "
		objBuilder.Append "'"&cfm_person1&"','"&cfm_person2&"','"&cfm_use&"','"&cfm_use_dept&"','"&cfm_comment&"', "
		objBuilder.Append "NOW(),'"&user_name&"')"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

'		dbconn.CommitTrans
'		dbconn.Close()
'	    Set dbconn = Nothing

 %>

	</table>
	</div>
</body>
</html>
