<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

order_no = request("order_no")
order_seq = request("order_seq")
order_date = request("order_date")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_order where (order_no = '"&order_no&"') and (order_seq = '"&order_seq&"') and (order_date = '"&order_date&"')"
Set Rs_order = DbConn.Execute(SQL)
if not Rs_order.eof then
    	order_no = Rs_order("order_no")
		order_seq = Rs_order("order_seq")
		order_date = Rs_order("order_date")
		order_buy_no = Rs_order("order_buy_no")
		order_buy_seq = Rs_order("order_buy_seq")
		order_buy_date = Rs_order("order_buy_date")
		
		order_goods_type = Rs_order("order_goods_type")
		order_company = Rs_order("order_company")
	    order_bonbu = Rs_order("order_bonbu")
		order_saupbu = Rs_order("order_saupbu")
		order_team = Rs_order("order_team")
	    order_org_code = Rs_order("order_org_code")
	    order_org_name = Rs_order("order_org_name")
	    order_emp_no = Rs_order("order_emp_no")
	    order_emp_name = Rs_order("order_emp_name")
		
	    order_bill_collect = Rs_order("order_bill_collect")
        order_collect_due_date = Rs_order("order_collect_due_date")
	    order_trade_no = Rs_order("order_trade_no")
        order_trade_name = Rs_order("order_trade_name")
        order_trade_person = Rs_order("order_trade_person")
		order_trade_email = Rs_order("order_trade_email")
		
        buy_out_method = ""
        buy_out_request_date = ""
		
		order_in_date = Rs_order("order_in_date")
        order_stock_company = Rs_order("order_stock_company")
        order_stock_code = Rs_order("order_stock_code")
        order_stock_name = Rs_order("order_stock_name")
		
        order_price = Rs_order("order_price")
        order_cost = Rs_order("order_cost")
        order_cost_vat = Rs_order("order_cost_vat")
		
        order_memo = Rs_order("order_memo")
        if order_memo = "" or isnull(order_memo) then
	           order_memo = Rs_order("order_memo")
           else
	           order_memo = replace(order_memo,chr(10),"<br>")
        end if
        order_ing = Rs_order("order_ing")

	    if order_collect_due_date = "0000-00-00" then
	          order_collect_due_date = ""
	    end if
		if order_in_date = "0000-00-00" then
	      order_in_date = ""
	    end if
   else
		order_buy_no = ""
		order_buy_seq = ""
		order_buy_date = ""
		order_goods_type = ""
		order_company = ""
	    order_bonbu = ""
		order_saupbu = ""
		order_team = ""
	    order_org_code = ""
	    order_org_name = ""
	    order_emp_no = ""
	    order_emp_name = ""
	    order_bill_collect = ""
        order_collect_due_date = ""
	    order_trade_no = ""
        order_trade_name = ""
        order_trade_person = ""
		order_trade_email = ""
        buy_out_method = ""
        buy_out_request_date = ""
		order_in_date = ""
        order_stock_company = ""
        order_stock_code = ""
        order_stock_name = ""
        order_price = 0
        order_cost = 0
        order_cost_vat = 0
        order_memo = ""
        order_ing = ""
end if
Rs_order.close()


if order_company = "���̿��������" then
      company_name = "(��)" + "���̿��������"
	  owner_name = "�����"
	  addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	  trade_no = "107-81-54150"
	  tel_no = "02) 853-5250"
	  e_mail = "js10547@k-won.co.kr"
   elseif order_company = "�޵�" then
              company_name = "(��)" + "�޵�"
			  owner_name = "������"
	          addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	          trade_no = "107-81-54150"
	          tel_no = "02) 853-5250"
	          e_mail = "js10547@k-won.co.kr"
		  elseif order_company = "���̳�Ʈ����" then
                     company_name = "���̳�Ʈ����" + "(��)"
					 owner_name = "���߿�"
	                 addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	                 trade_no = "107-81-54150"
	                 tel_no = "02) 853-5250"
	                 e_mail = "js10547@k-won.co.kr"
				 elseif order_company = "����������ġ" then
                        company_name = "(��)" + "����������ġ"	
						owner_name = "�ڹ̾�"
	                    addr_name = "����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)"
	                    trade_no = "119-86-78709"
	                    tel_no = "02) 6116-8248"
	                    e_mail = "pshwork27@k-won.co.kr"
end if 

sql = "select * from met_order_goods where (og_order_no = '"&order_no&"') and (og_order_seq = '"&order_seq&"') and (og_order_date = '"&order_date&"') ORDER BY og_seq,og_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "�� �� ��"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>������� �ý���</title>
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
                factory.printing.portrait = true; //��¹��� ����: true - ����, false - ����
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
				
					document.frm.method = "post";
//					document.frm.enctype = "multipart/form-data";
					document.frm.action = "met_buy_order_prt_ok.asp";
					document.frm.submit();
            }
        </script>
        <style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style14L {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: left; }
		.style14C {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style14R {font-size: 14px; font-family: "����ü", "����ü", Seoul; text-align: right; }
		.style18L {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "����ü", "����ü", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "����ü", "����ü", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
-->
    </style>
	</head>
	<style media="print"> 
    .noprint     { display: none }
    </style>
	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="wrap">			
			<div id="container">
				<form action="met_buy_order_print.asp" method="post" name="frm">
				<div class="gView">
				<table width="1150" cellpadding="0" cellspacing="0">
				  <tr>
				    <td height="50px" class="style32BC"><strong><%=title_line%></strong></td>
			      </tr>
				  </table>
					<br>
				<table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
							<col width="20%" >
							<col width="30%" >
							<col width="4%" >
							<col width="16%" >
							<col width="30%" >
						</colgroup>
						<thead>
							<tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">��������(��ȣ)</td>
                              <td align="center" class="style14C"><%=order_date%>&nbsp;(<%=order_no%>&nbsp;<%=order_seq%>)</td>
                              <th rowspan="6" align="center" class="style14C" style="background:#f8f8f8;">��<br>��<br>ó</th>
                              <th align="center" class="style14C" style="background:#f8f8f8;">����ڵ�Ϲ�ȣ</th>
                              <td align="center" class="style14C"><%=trade_no%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">�ŷ�ó��</td>
                              <td align="center" class="style14C"><%=order_trade_name%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">��ȣ</th>
                              <td align="center" class="style14C"><%=company_name%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">TEL No.</td>
                              <td align="center" class="style14C"><%=tel_no%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">�ּ�</th>
                              <td class="left"><font style="font-size:14px"><%=addr_name%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">FAX No.</td>
                              <td align="center" class="style14C"><%=tel_no%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">���ִ����</th>
                              <td align="center" class="style14C"><%=order_emp_name%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">�����</td>
                              <td align="center" class="style14C"><%=order_trade_person%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">TEL No.</th>
                              <td align="center" class="style14C"><%=tel_no%></td>
						    </tr>
                            <tr>
                              <td height="20" align="center" class="style14C" style="background:#f8f8f8;">������</td>
                              <td align="center" class="style14C"><%=order_in_date%></td>
                              <th align="center" class="style14C" style="background:#f8f8f8;">FAX No.</th>
                              <td align="center" class="style14C"><%=tel_no%></td>
						    </tr>
						</thead>
					</table>
                     <br>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="*" >
                              <col width="20%" >
                              <col width="10%" >
                              <col width="10%" >
							  <col width="12%" >
							  <col width="14%" >
							  <col width="10%" >
                        </colgroup>
						 <thead>
                              <tr bgcolor="#f8f8f8">
                                <th class="first" height="30" align="center" scope="col" class="style14C">ǰ ��</th>
                                <th scope="col" align="center" class="style14C">�� ��</th>
                                <th scope="col" align="center" class="style14C">����</th>
                                <th scope="col" align="center" class="style14C">����</th>
                                <th scope="col" align="center" class="style14C">�� ��</th>
                                <th scope="col" align="center" class="style14C">�� ��</th>
                                <th scope="col" align="center" class="style14C">���</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						do until rs.eof or rs.bof
                             
	           		 %>
							<tr>
                                <td height="30" align="center" class="style14C"><%=rs("og_goods_name")%>&nbsp;</td>
                                <td align="center" class="style14C"><%=rs("og_standard")%>&nbsp;</td>
                                <td align="center" class="style14C">&nbsp;</td>
                                <td align="right" class="style14R"><%=formatnumber(rs("og_qty"),0)%>&nbsp;</td>
                                <td align="right" class="style14R"><%=formatnumber(rs("og_unit_cost"),0)%>&nbsp;</td>
                                <td align="right" class="style14R" ><%=formatnumber(rs("og_amt"),0)%>&nbsp;</td>
                                <td align="center" class="style14C">&nbsp;</td>
							</tr>
					<%
							rs.movenext()
						loop
						rs.close()
					%>
                            <tr>
                                <td height="30" align="center" class="style14C" style="background:#f8f8f8;">���ް���</td>
                                <td align="right" class="style14R"><%=formatnumber(order_cost,0)%>&nbsp;</td>
                                <td colspan="2" align="center" class="style14C" style="background:#f8f8f8;">�ΰ�����</td>
                                <td align="right" class="style14R"><%=formatnumber(order_cost_vat,0)%>&nbsp;</td>
                                <td align="center" class="style14C" style="background:#f8f8f8;">�հ�ݾ�</td>
                                <td align="right" class="style14R"><%=formatnumber(order_price,0)%>&nbsp;</td>
							</tr>
						</tbody>
					</table> 
                    <br>
                    <h3 class="stit">1. �ͻ��� ���� ��â�Ͻ��� ����մϴ�.</h3>
                    <h3 class="stit">&nbsp;&nbsp;&nbsp;���� ���� �����Ͽ��� �������� �ؼ��Ͽ� �԰� �ٶ��ϴ�.</h3> 
                    <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
							<col width="20%" >
							<col width="80%" >
						</colgroup>
						<thead>
							<tr>
                              <td height="30" align="center" class="style14C" style="background:#f8f8f8;">��ݰ��� ����</td>
                              <td class="left" ><font style="font-size:14px">&nbsp;<%=order_collect_due_date%>&nbsp;-&nbsp;<%=order_bill_collect%></td>
						    </tr>
                            <tr>
                              <td height="30" align="center" class="style14C" style="background:#f8f8f8;">��ǰ ���</td>
                              <td class="left" ><font style="font-size:14px">&nbsp;<%=order_stock_name%>&nbsp;-&nbsp;<%=addr_name%></td>
						    </tr>
                            <tr>
                              <td height="30" align="center" class="style14C" style="background:#f8f8f8;">Ư�� ����</td>
                              <td class="left"><font style="font-size:14px">&nbsp;<%=order_memo%></td>
						    </tr>
						</thead>
					</table>  
				<table width="1150" border="0" cellpadding="0" cellspacing="0" align="center" class="onlyprint">    
				  <tr>
				     <td colspan="2" height="100" align="center"><font style="font-size:16px"><strong>�� ����� ���� ���� �մϴ�.</td>
	              </tr>
	              <tr>
		             <td colspan="2" height="60" align="right" width="100%"><font style="font-size:14px"><%=mid(cstr(now()),1,4)%>��&nbsp;<%=mid(cstr(now()),6,2)%>��&nbsp;<%=mid(cstr(now()),9,2)%>��<br/><br/>
		����� ��õ�� ���������2�� 18(�븢��ũ��Ÿ�� 1�� 6��)</td>
	             </tr>
	             <tr>  
	                <td height="60" align="right" width="95%"><font style="font-size:14px"><br><br>�ֽ�ȸ�� ���̿��������<br/>
		<font style="font-size:14px">��ǥ�̻� </font><font style="font-size:16px"><b>�����</b></font></td>
                    <td height="60" align="right" valign="middle" width="5%"><img src="image/k-won001.png" width=80 height=80 alt="" align=right></td>
	             </tr>                    
				</table>
                <br><br><br>
				<table width="1150" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<br>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>            
                    		<span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>            
                    	</div>
    				</div>
				    <br>                 
                    </td>
			      </tr>
				</table>
                <input type="hidden" name="old_order_no" value="<%=order_no%>">
				<input type="hidden" name="old_order_seq" value="<%=order_seq%>">
                <input type="hidden" name="old_order_date" value="<%=order_date%>">
                
                <input type="hidden" name="order_buy_no" value="<%=order_buy_no%>">
				<input type="hidden" name="order_buy_seq" value="<%=order_buy_seq%>">
                <input type="hidden" name="order_buy_date" value="<%=order_buy_date%>">
			</form>
		</div>				
	</div>        				
	</body>
</html>

