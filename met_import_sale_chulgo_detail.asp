<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

stock_goods_code = request("stock_goods_code")
stock_goods_type = request("stock_goods_type")
stock_code = request("stock_code")
stock_name = request("stock_name")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
Set rs = DbConn.Execute(SQL)
if not rs.eof then
    	goods_code = rs("goods_code")
		goods_grade = rs("goods_grade")
        goods_gubun = rs("goods_gubun")
	    goods_name = rs("goods_name")
	    goods_standard = rs("goods_standard")
	    goods_type = rs("goods_type")
		goods_model = rs("goods_model")
		part_number = rs("part_number")
   else
		goods_code = ""
		goods_grade = ""
        goods_gubun = ""
	    goods_name = ""
	    goods_standard = ""
	    goods_type = ""
		goods_model = ""
		part_number = ""
end if
rs.close()

sql = "select * from met_chulgo_goods where (cg_goods_code = '"&stock_goods_code&"') and (cg_goods_type = '"&stock_goods_type&"') and (chulgo_stock = '"&stock_code&"') ORDER BY chulgo_date DESC"
Rs.Open Sql, Dbconn, 1

title_line = goods_name + " ǰ�� < " + stock_name + " >â�� �����Ȳ(N/W)"

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
							    <th>â��</th>
							    <td class="left" colspan="3"><%=stock_name%>&nbsp;(<%=stock_code%>)</td>
                                <th>ǰ�񱸺�</th>
							    <td class="left"><%=goods_gubun%>&nbsp;</td>
						    </tr>
                            <tr>
                                <th>ǰ���ڵ�</th>
							    <td class="left"><%=goods_code%>&nbsp;</td>
							    <th>ǰ���</th>
							    <td class="left"><%=goods_name%>&nbsp;</td>
							    <th>����</th>
							    <td class="left"><%=goods_grade%>&nbsp;</td>
 							</tr>
                            <tr>
							    <th>�԰�</th>
							    <td class="left"><%=goods_standard%>&nbsp;</td>
                                <th>��</th>
							    <td class="left"><%=goods_model%>&nbsp;</td>
                                <th>Part_No.</th>
							    <td class="left"><%=part_number%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� ���ں� ��� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="10%" >
                            <col width="12%" >
                            <col width="8%" >
                            <col width="10%" >
                            <col width="*" >
                            <col width="14%" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col">�������</th>
                                <th scope="col">�뵵����</th>
                                <th scope="col">�����</th>
                                <th scope="col">����ȣ</th>
                                <th scope="col">��ǥ No.</th>
                                <th scope="col">������</th>
                                <th scope="col">���ݾ�</th>
                                <th scope="col">����</th>
                                <th scope="col">������</th>
                                
							</tr>
						</thead>
						<tbody>     
						<%
							i = 0
							h_go_qty = 0
							h_go_amt = 0
							do until rs.eof or rs.bof
							     i = i + 1
							     yymmdd = mid(cstr(rs("chulgo_date")),3,2) + mid(cstr(rs("chulgo_date")),6,2)  + mid(cstr(rs("chulgo_date")),9,2)
							     chulgo_no = yymmdd + rs("chulgo_stock")
								 
								 if rs("cg_qty") > 0 then
									 h_go_qty = h_go_qty + rs("cg_qty")
									 h_go_amt = h_go_amt + rs("cg_amt")
						%>
							<tr>
                                <td><%=rs("chulgo_date")%>&nbsp;</td>
                                <td><%=rs("cg_goods_type")%>&nbsp;</td>
                                <td><%=rs("cg_type")%>&nbsp;</td>
                                <td><%=chulgo_no%>&nbsp;<%=rs("chulgo_seq")%></td>
                                <td><%=rs("rl_service_no")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("cg_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("cg_amt"),0)%>&nbsp;</td>
                                <td><%=rs("rl_trade_name")%>&nbsp;</td>
                                <td><%=rs("rl_trade_dept")%>&nbsp;</td>
							</tr>
						<%
								end if
								rs.movenext()
							loop
							rs.close()
							
		' �������� �԰�������, â���̵� �԰�� ..... �԰���Ȳ�� ������ �Ұ�										
							
							
						%>
                            <tr>
                                <td colspan="5" style="background:#ffe8e8;">�� ��</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_go_qty,0)%>&nbsp;</td>
                                <td class="right" style="background:#ffe8e8;"><%=formatnumber(h_go_amt,0)%>&nbsp;</td>
								<td colspan="2" style="background:#ffe8e8;">&nbsp;</td>
							</tr>
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
                    <input type="hidden" name="stock_goods_code" value="<%=stock_goods_code%>">
					<input type="hidden" name="stock_goods_type" value="<%=stock_goods_type%>">
	     </form>
    	</div>				
	  </div>     
	</body>
</html>

