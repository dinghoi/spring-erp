<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

card_no = request("card_no")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs_hist = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from card_owner INNER JOIN memb ON card_owner.emp_no = memb.emp_no where card_no ='" + card_no + "'"
'response.write(sql)
set rs = dbconn.execute(sql)

sql = "select * from card_owner_history where card_no = '" + card_no + "' order by history_seq desc"
'response.write(sql)
rs_hist.Open Sql, Dbconn, 1

title_line = "ī�� ����� ���� History ��ȸ"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ȸ�� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
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
							  <th>ī������</th>
							  <td class="left"><%=rs("card_type")%></td>
							  <th>ī���ȣ</th>
							  <td class="left"><%=rs("card_no")%></td>
					      	</tr>
							<tr>
							  <th>���౸��</th>
							  <td class="left"><%=rs("card_issue")%>&nbsp;</td>
							  <th>����ѵ�</th>
							  <td class="left"><%=rs("card_limit")%>&nbsp;</td>
					      	</tr>
							<tr>
							  <th>��ȿ�Ⱓ</th>
							  <td class="left"><%=rs("valid_thru")%>&nbsp;</td>
							  <th>�߱���</th>
							  <td class="left"><%=rs("create_date")%>&nbsp;</td>
					      	</tr>
						</tbody>
					</table>
					<h3 class="stit">* History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="15%" >
							<col width="20%" >
							<col width="10%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">�μ�</th>
								<th scope="col">������</th>
								<th scope="col">������</th>
								<th scope="col" class="left">��������</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td class="first">00</td>
								<td><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("start_date")%></td>
								<td>&nbsp;</td>
								<td class="left">���� �����</td>
							</tr>
						<%
                        do until rs_hist.eof 
                        %>
							<tr>
								<td class="first"><%=rs_hist("history_seq")%></td>
								<td><%=rs_hist("emp_name")%>&nbsp;<%=rs_hist("emp_job")%></td>
								<td><%=rs_hist("org_name")%></td>
								<td><%=rs_hist("start_date")%></td>
								<td><%=rs_hist("end_date")%></td>
								<td class="left"><%=rs_hist("mod_memo")%></td>
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

