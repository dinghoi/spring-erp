<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

buy_no = request("buy_no")
buy_date = request("buy_date")
buy_seq = request("buy_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_buy where (buy_no = '"&buy_no&"') and (buy_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"')"
Set Rs_buy = DbConn.Execute(SQL)
if not Rs_buy.eof then
    	buy_no = Rs_buy("buy_no")
		buy_date = Rs_buy("buy_date")
		buy_goods_type = Rs_buy("buy_goods_type")
		buy_company = Rs_buy("buy_company")
	    buy_bonbu = Rs_buy("buy_bonbu")
		buy_saupbu = Rs_buy("buy_saupbu")
		buy_team = Rs_buy("buy_team")
	    buy_org_code = Rs_buy("buy_org_code")
	    buy_org_name = Rs_buy("buy_org_name")
	    buy_emp_no = Rs_buy("buy_emp_no")
	    buy_emp_name = Rs_buy("buy_emp_name")
	    buy_bill_collect = Rs_buy("buy_bill_collect")
        buy_collect_due_date = Rs_buy("buy_collect_due_date")
	    buy_trade_no = Rs_buy("buy_trade_no")
        buy_trade_name = Rs_buy("buy_trade_name")
        buy_trade_person = Rs_buy("buy_trade_person")
		buy_trade_email = Rs_buy("buy_trade_email")
        buy_out_method = Rs_buy("buy_out_method")
        buy_out_request_date = Rs_buy("buy_out_request_date")
        buy_price = Rs_buy("buy_price")
        buy_cost = Rs_buy("buy_cost")
        buy_cost_vat = Rs_buy("buy_cost_vat")
        buy_memo = Rs_buy("buy_memo")
        if buy_memo = "" or isnull(buy_memo) then
	           buy_memo = Rs_buy("buy_memo")
           else
	           buy_memo = replace(buy_memo,chr(10),"<br>")
        end if
        buy_ing = Rs_buy("buy_ing")
		buy_sign_yn = Rs_buy("buy_sign_yn")
	    buy_sign_no = Rs_buy("buy_sign_no")
	    buy_sign_date = Rs_buy("buy_sign_date")
		buy_att_file = Rs_buy("buy_att_file")

	    if buy_out_request_date = "0000-00-00" then
	          buy_out_request_date = ""
	    end if
   else
		buy_company = ""
	    buy_bonbu = ""
		buy_saupbu = ""
		buy_team = ""
	    buy_org_code = ""
	    buy_org_name = ""
	    buy_emp_no = ""
	    buy_emp_name = ""
	    buy_bill_collect = ""
        buy_collect_due_date = ""
	    buy_trade_no = ""
        buy_trade_name = ""
        buy_trade_person = ""
		buy_trade_email = ""
        buy_out_method = ""
        buy_out_request_date = ""
        buy_price = 0
        buy_cost = 0
        buy_cost_vat = 0
        buy_memo = ""
        buy_ing = ""
		buy_att_file = ""
end if
Rs_buy.close()

sql = "select * from met_buy_goods where (bg_no = '"&buy_no&"') and (bg_date = '"&buy_date&"') and (buy_seq = '"&buy_seq&"') ORDER BY bg_seq,bg_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "����ǰ�� ��ȸ"

view_att_file = buy_att_file
path = "/met_upload"


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
				a=confirm('����ǰ�Ǹ� ����ϰڽ��ϱ�?')
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
//			function approve_request(slip_id,slip_no,slip_seq) 
			function approve_request() 
				{
				a=confirm('���� ��û�Ͻðڽ��ϱ�?')
				if (a==true) {
//					document.frm.action = "met_buy_approve_ok.asp?slip_id="+slip_id+'&slip_no='+slip_no+'&slip_seq='+slip_seq;
					document.frm.action = "met_buy_approve_ok.asp";
					document.frm.submit();
				}
				return false;
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
                                <th>���Ź�ȣ</th>
							    <td class="left"><%=buy_no%>&nbsp;<%=buy_seq%></td>
							    <th>��������</th>
							    <td class="left"><%=buy_goods_type%></td>
							    <th>��������</th>
							    <td class="left"><%=buy_date%></td>
 							</tr>
                            <tr>
							    <th>����ȸ��</th>
							    <td class="left"><%=buy_company%></td>
							    <th>�����</th>
							    <td class="left"><%=buy_saupbu%></td>
							    <th>���Ŵ��</th>
							    <td class="left"><%=buy_org_name%>&nbsp;<%=buy_emp_name%></td>
						    </tr>
							<tr>
                                <th>����ó</th>
							    <td class="left"><%=buy_trade_name%></td>
							    <th>����ڹ�ȣ</th>
							    <td class="left"><%=buy_trade_no%></td>
							    <th>�����</th>
							    <td class="left"><%=buy_trade_person%></td>
						    </tr>
                            <tr>
                                <th>�̸���</th>
							    <td class="left"><%=buy_trade_email%></td>
							    <th>���<br>���޹��</th>
							    <td class="left"><%=buy_bill_collect%></td>
							    <th>���޿�����</th>
							    <td class="left"><%=buy_collect_due_date%></td>
						    </tr>
							<tr>
							  <th>���</th>
							  <td colspan="5" class="left"><%=buy_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� ���� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="16%" >
							<col width="14%" >
							<col width="8%" >
							<col width="12%" >
							<col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�뵵����</th>
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�԰�</th>
								<th scope="col">����</th>
								<th scope="col">���Դܰ�</th>
								<th scope="col">���Աݾ�</th>
							</tr>
						</thead>
						<tbody>     
						<%
							buy_cost_tot = 0
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
							
							     buy_hap = rs("bg_qty") * rs("bg_unit_cost")
							     buy_cost_tot = buy_cost_tot + buy_hap
							
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("bg_goods_type")%>&nbsp;</td>
								<td><%=rs("bg_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("bg_goods_code")%>&nbsp;</td>
                                <td><%=rs("bg_goods_name")%>&nbsp;</td>
                                <td><%=rs("bg_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("bg_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("bg_unit_cost"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(buy_hap,0)%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						</tbody>
					</table>
                    <br>
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
                        <% 
						    buy_vat_hap = int(buy_cost_tot * (10 / 100))
							buy_tot_price = buy_cost_tot + buy_vat_hap
						%>
							<tr>
							  <th>�����Ѿ�</th>
							  <td class="right"><%=formatnumber(buy_tot_price,0)%></td>
							  <th>���űݾ�</th>
							  <td class="right"><%=formatnumber(buy_cost_tot,0)%></td>
							  <th>�ΰ���</th>
							  <td class="right"><%=formatnumber(buy_vat_hap,0)%></td>
						    </tr>
							<tr>
							  <th>÷��</th>
							  <td colspan="5" class="left">
                        <% 
                           If buy_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=buy_att_file%>"><%=buy_att_file%></a>
                        <%    Else %>
				                    &nbsp;
                        <% 
						   End If %>
                              </td>
						    </tr>
						</tbody>
					</table>
          	     <br>
     				<div class="noprint">
                        <div align=center>
                        <% if buy_sign_yn = "N" then	%>
                            <span class="btnType01"><input type="button" value="�����û" onclick="javascript:approve_request('<%=buy_no%>','<%=buy_seq%>','<%=buy_date%>');"></span>
                        <% end if	%>
                        <% if cancel_yn = "Y" then	%>
                            <span class="btnType01"><input type="button" value="��ǥ���" onclick="javascript:frmcheck();"></span>
                        <% end if	%>
                            <span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>
                            <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="buy_no" value="<%=buy_no%>">
					<input type="hidden" name="buy_date" value="<%=buy_date%>">
					<input type="hidden" name="buy_seq" value="<%=buy_seq%>">
					<input type="hidden" name="cancel_yn" value="<%=cancel_yn%>">      				
	     </form>
    	</div>				
	  </div>     
	</body>
</html>

