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
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_mv_go where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')"
Set Rs_chul = DbConn.Execute(SQL)
if not Rs_chul.eof then
    	rele_stock = Rs_chul("rele_stock")
        rele_seq = Rs_chul("rele_seq")
	    rele_date = Rs_chul("rele_date")
		
        chulgo_id = Rs_chul("chulgo_id")
        chulgo_type = Rs_chul("chulgo_type")
		chulgo_goods_type = Rs_chul("chulgo_goods_type")
		chulgo_stock_company = Rs_chul("chulgo_stock_company")
        chulgo_stock_name = Rs_chul("chulgo_stock_name")
        chulgo_emp_no = Rs_chul("chulgo_emp_no")
        chulgo_emp_name = Rs_chul("chulgo_emp_name")
        chulgo_company = Rs_chul("chulgo_company")
        chulgo_bonbu = Rs_chul("chulgo_bonbu")
        chulgo_saupbu = Rs_chul("chulgo_saupbu")
        chulgo_team = Rs_chul("chulgo_team")
        chulgo_org_name = Rs_chul("chulgo_org_name")

        in_stock_date = Rs_chul("in_stock_date")
		chulgo_memo = Rs_chul("chulgo_memo")
	    if in_stock_date = "0000-00-00" then
	          in_stock_date = ""
	    end if
   else
		rele_stock = ""
        rele_seq = ""
	    rele_date = ""
        chulgo_id = ""
        chulgo_type = ""
		chulgo_stock_company = ""
        chulgo_stock_name = ""
        chulgo_emp_no = ""
        chulgo_emp_name = ""
        chulgo_company = ""
        chulgo_bonbu = ""
        chulgo_saupbu = ""
        chulgo_team = ""
        chulgo_org_name = ""

        in_stock_date = ""
		chulgo_memo = ""
end if
Rs_chul.close()

sql = "select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')"
Set Rs_reg = DbConn.Execute(SQL)
if not Rs_reg.eof then
    	rele_stock_company = Rs_reg("rele_stock_company")
        rele_stock_name = Rs_reg("rele_stock_name")
        rele_emp_no = Rs_reg("rele_emp_no")
        rele_emp_name = Rs_reg("rele_emp_name")
        rele_company = Rs_reg("rele_company")
        rele_bonbu = Rs_reg("rele_bonbu")
        rele_saupbu = Rs_reg("rele_saupbu")
        rele_team = Rs_reg("rele_team")
        rele_org_name = Rs_reg("rele_org_name")

        chulgo_rele_date = Rs_reg("chulgo_rele_date")
   else
		rele_stock_company = ""
        rele_stock_name = ""
        rele_emp_no = ""
        rele_emp_name = ""
        rele_company = ""
        rele_bonbu = ""
        rele_saupbu = ""
        rele_team = ""
        rele_org_name = ""

        chulgo_rele_date = ""
end if
Rs_reg.close()

sql = "select * from met_mv_go_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')  ORDER BY cg_goods_seq,cg_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "â���̵� ��� ��ȸ"

view_att_file = rele_att_file
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
				a=confirm('â���̵� ����Ƿڸ� ����ϰڽ��ϱ�?')
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
					document.frm.action = "met_move_chulgo_approve_ok.asp";
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
				<form method="post" name="frm" action="met_move_chulgo_cancel.asp">
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
                                <th>�������</th>
							    <td class="left"><%=chulgo_date%></td>
							    <th>�����</th>
							    <td class="left"><%=chulgo_emp_name%>&nbsp;(<%=chulgo_org_name%>)</td>
							    <th>���â��</th>
							    <td class="left"><%=chulgo_stock_name%>&nbsp;(<%=chulgo_stock_company%>)</td>
 							</tr>
                            <tr>
                                <th>��û����</th>
							    <td class="left"><%=rele_date%></td>
							    <th>�뵵����</th>
							    <td class="left"><%=chulgo_goods_type%></td>
							    <th>��ûâ��</th>
							    <td class="left"><%=rele_stock_name%>&nbsp;(<%=rele_stock_company%>)</td>
 							</tr>
                            <tr>
							    <th>��û��</th>
							    <td class="left"><%=rele_emp_name%>&nbsp;(<%=rele_org_name%>)</td>
							    <th>����û��</th>
							    <td class="left"><%=chulgo_rele_date%></td>
							    <th>��ûâ��<br>�԰���</th>
							    <td class="left"><%=in_stock_date%>&nbsp;</td>
						    </tr>
							<tr>
							  <th>���</th>
							  <td colspan="5" class="left"><%=chulgo_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� â���̵� ��� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="8%" >
                            <col width="*" >
                            <col width="12%" >
							<col width="15%" >
							<col width="15%" >
							<col width="8%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�뵵����</th>
                                <th scope="col">����</th>
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�԰�</th>
								<th scope="col">�Ƿڼ���</th>
                                <th scope="col">������</th>
							</tr>
						</thead>
						<tbody>     
						<%
							i = 0
							do until rs.eof or rs.bof
							     i = i + 1
					
						%>
							<tr>
								<td class="first"><%=i%></td>
                                <td><%=rs("cg_goods_type")%>&nbsp;</td>
                                <td><%=rs("cg_goods_grade")%>&nbsp;</td>
								<td><%=rs("cg_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("cg_goods_code")%>&nbsp;</td>
                                <td><%=rs("cg_goods_name")%>&nbsp;</td>
                                <td><%=rs("cg_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rl_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("cg_qty"),0)%>&nbsp;</td>
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
                            <span class="btnType01"><input type="button" value="����� ���" onClick="pop_Window('met_move_chulgo_referral_print.asp?chulgo_date=<%=chulgo_date%>&chulgo_stock=<%=chulgo_stock%>&chulgo_seq=<%=chulgo_seq%>','met_move_chulgo_print_pop','scrollbars=yes,width=750,height=600')"></span>
                            <span class="btnType01"><input type="button" value="�μ��� ���" onClick="pop_Window('met_move_chulgo_receip_print.asp?chulgo_date=<%=chulgo_date%>&chulgo_stock=<%=chulgo_stock%>&chulgo_seq=<%=chulgo_seq%>','met_move_chulgo_receip_pop','scrollbars=yes,width=750,height=600')"></span>
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="rele_stock" value="<%=rele_stock%>">
					<input type="hidden" name="rele_seq" value="<%=rele_seq%>">
					<input type="hidden" name="rele_date" value="<%=rele_date%>">
					<input type="hidden" name="cancel_yn" value="<%=cancel_yn%>">      				
	     </form>
    	</div>				
	  </div>     
	</body>
</html>
