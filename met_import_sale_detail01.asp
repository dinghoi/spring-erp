<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

chulgo_date = request("chulgo_date")
chulgo_stock = request("chulgo_stock")
chulgo_seq = request("chulgo_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_chulgo where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')"
Set rs = DbConn.Execute(SQL)
if not rs.eof then
    	chulgo_goods_type = rs("chulgo_goods_type")
        chulgo_id = rs("chulgo_id")
	    service_no = rs("service_no")
	    chulgo_trade_name = rs("chulgo_trade_name")
	    chulgo_trade_dept = rs("chulgo_trade_dept")
	    chulgo_type = rs("chulgo_type")
	    service_no = rs("service_no")
	
        chulgo_stock_company = rs("chulgo_stock_company")
        chulgo_stock_name = rs("chulgo_stock_name")
        chulgo_emp_no = rs("chulgo_emp_no")
        chulgo_emp_name = rs("chulgo_emp_name")
        chulgo_company = rs("chulgo_company")
        chulgo_bonbu = rs("chulgo_bonbu")
        chulgo_saupbu = rs("chulgo_saupbu")
        chulgo_team = rs("chulgo_team")
        chulgo_org_name = rs("chulgo_org_name")
        chulgo_memo = rs("chulgo_memo")
		rele_no = rs("rele_no")
		rele_seq = rs("rele_seq")
		rele_date = rs("rele_date")
   else
		chulgo_goods_type = ""
        chulgo_id = ""
	    service_no = ""
	    chulgo_trade_name = ""
	    chulgo_trade_dept = ""
	    chulgo_type = ""
	    service_no = ""
	
        chulgo_stock_company = ""
        chulgo_stock_name = ""
        chulgo_emp_no = ""
        chulgo_emp_name = ""
        chulgo_company = ""
        chulgo_bonbu = ""
        chulgo_saupbu = ""
        chulgo_team = ""
        chulgo_org_name = ""
        chulgo_memo = ""
end if
rs.close()

sql = "select * from met_chulgo_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"') ORDER BY cg_goods_seq,cg_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "N/W ���� ��� ����ȸ"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>������� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}		
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}					
			function chkfrm() {
						
				{
				a=confirm('��� ����ϰڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
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
				<h3 class="insa"><%=title_line%></h3>
				<form method="post" name="frm" action="met_buy_cancel.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="22%" >
							<col width="11%" >
							<col width="*" >
						</colgroup>
						<tbody> 
							<tr>
                                <th>ȸ��</th>
							    <td class="left"><%=chulgo_company%></td>
							    <th>�����</th>
							    <td class="left"><%=chulgo_saupbu%></td>
							    <th>���â��</th>
							    <td class="left">(<%=chulgo_stock_company%>)&nbsp;<%=chulgo_stock_name%></td>
 							</tr>
                            <tr>
							    <th>�������(��ȣ)</th>
							    <td class="left"><%=chulgo_date%>&nbsp;(<%=rele_no%>&nbsp;<%=rele_seq%>)</td>
							    <th>�뵵����</th>
							    <td class="left"><%=chulgo_goods_type%></td>
							    <th>�����</th>
							    <td class="left"><%=chulgo_org_name%>&nbsp;<%=chulgo_emp_name%></td>
						    </tr>
                            <tr>
							    <th>���񽺹�ȣ</th>
							    <td class="left"><%=service_no%></td>
							    <th>����</th>
							    <td class="left"><%=chulgo_trade_name%>&nbsp;(<%=chulgo_trade_dept%>)</td>
                                <th>�������</th>
							    <td class="left"><%=chulgo_id%></td>
						    </tr>
                            <tr>
							  <th>���</th>
							  <td colspan="5" class="left"><%=chulgo_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� ��� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
                            <col width="6%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="10%" >
							<col width="*" >
							<col width="12%" >
                            <col width="12%" >
							<col width="8%" >
                            <col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">����</th>
                                <th scope="col">�뵵����</th>
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�԰�</th>
                                <th scope="col">Part_No.</th>
                                <th scope="col">������</th>
                                <th scope="col" class="right">���ݾ�</th>
							</tr>
						</thead>
						<tbody>     
						<%
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
							     stock_goods_code = rs("cg_goods_code")
								 sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
                                 Set Rs_good = DbConn.Execute(SQL)
                                 if not Rs_good.eof then
    	                               goods_model = Rs_good("goods_model")
		                               part_number = Rs_good("part_number")
                                    else
		                               goods_model = ""
		                               part_number = ""
                                 end if
                                 Rs_good.close()
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("cg_goods_grade")%>&nbsp;</td>
                                <td><%=rs("cg_goods_type")%>&nbsp;</td>
								<td><%=rs("cg_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("cg_goods_code")%>&nbsp;</td>
                                <td><%=rs("cg_goods_name")%>&nbsp;</td>
                                <td><%=rs("cg_standard")%>&nbsp;</td>
                                <td><%=part_number%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("cg_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("cg_amt"),0)%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						</tbody>
					</table>
          	     <br>
     				<div class="noprint">
                        <div align=center>
                            <span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>
                            <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="chulgo_date" value="<%=chulgo_date%>">
					<input type="hidden" name="chulgo_stock" value="<%=chulgo_stock%>">
					<input type="hidden" name="chulgo_seq" value="<%=chulgo_seq%>">
	     </form>
    	</div>				
	  </div>     
	</body>
</html>

