<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

company = request("company")
dept_code = request("dept_code")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_hist = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from asset_dept where company = '" + company + "' and dept_code = '" + dept_code + "'"
set rs = dbconn.execute(sql)

etc_code = "75" + cstr(company)
sql = "select * from etc_code where etc_code = '" + etc_code + "'"
set rs_etc = dbconn.execute(sql)

sql = "select * from asset where company = '" + company + "' and dept_code = '" + dept_code + "' and inst_process = 'Y' order by gubun asc"
rs_hist.Open Sql, Dbconn, 1

title_line = "������ �ڻ� ���γ���"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}

			function printWindow(){
        //		viewOff("button");   
                factory.printing.header = ""; //�Ӹ��� ����
                factory.printing.footer = ""; //������ ����
                factory.printing.portrait = false; //��¹��� ����: true - ����, false - ����
                factory.printing.leftMargin = 13; //���� ���� ����
                factory.printing.topMargin = 10; //���� ���� ����
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

	</head>

	<style media="print"> 
    .noprint     { display: none }
    </style>

	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>ȸ��</th>
							  <td class="left"><%=rs_etc("etc_name")%></td>
							  <th>��������</th>
							  <td class="left"><%=rs("high_org")%></td>
					      	</tr>
							<tr>
							  <th>������</th>
							  <td class="left" colspan="3"><%=rs("org_first")%>&nbsp;/&nbsp;<%=rs("org_second")%>&nbsp;/&nbsp;<%=rs("dept_name")%></td>
					      	</tr>
							<tr>
							  <th>�����</th>
							  <td class="left"><%=rs("person")%></td>
							  <th>��ȭ��ȣ</th>
							  <td class="left"><%=rs("tel_ddd")%>&nbsp;-&nbsp;<%=rs("tel_no1")%>&nbsp;-&nbsp;<%=rs("tel_no2")%></td>
					      	</tr>
						</tbody>
					</table>
					<h3 class="stit">* �ڻ� ����</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="20%" >
							<col width="15%" >
							<col width="*" >
							<col width="10%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�ڻ��ȣ</th>
								<th scope="col">�ڻ��ڵ�</th>
								<th scope="col">�ڻ��</th>
								<th scope="col">�����</th>
								<th scope="col">��ġ��</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs_hist.eof
							i = i + 1			
                        %>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=mid(rs_hist("asset_no"),1,2)%>-<%=mid(rs_hist("asset_no"),3,6)%>-<%=right(rs_hist("asset_no"),4)%></td>
								<td><%=rs_hist("company")%>-<%=rs_hist("gubun")%>-<%=rs_hist("code_seq")%></td>
								<td><%=rs_hist("asset_name")%></td>
								<td><%=rs_hist("user_name")%></td>
								<td><%=rs_hist("install_date")%></td>
							</tr>
						<%
                            rs_hist.movenext()  
                        loop
                        rs_hist.Close()
                        %>
						</tbody>
					</table>                    
					<br>
				</form>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				</div>
			</div>
	</body>
</html>

