<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

rele_date = request("rele_date")
rele_stock = request("rele_stock")
rele_seq = request("rele_seq")

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

sql = "select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')"
Set Rs_chul = DbConn.Execute(SQL)
if not Rs_chul.eof then
    	rele_stock = Rs_chul("rele_stock")
        rele_seq = Rs_chul("rele_seq")
	    rele_date = Rs_chul("rele_date")
        rele_id = Rs_chul("rele_id")
        rele_goods_type = Rs_chul("rele_goods_type")
		rele_stock_company = Rs_chul("rele_stock_company")
        rele_stock_name = Rs_chul("rele_stock_name")
        rele_emp_no = Rs_chul("rele_emp_no")
        rele_emp_name = Rs_chul("rele_emp_name")
        rele_company = Rs_chul("rele_company")
        rele_bonbu = Rs_chul("rele_bonbu")
        rele_saupbu = Rs_chul("rele_saupbu")
        rele_team = Rs_chul("rele_team")
        rele_org_name = Rs_chul("rele_org_name")

        chulgo_rele_date = Rs_chul("chulgo_rele_date")
		chulgo_ing = Rs_chul("chulgo_ing")
        chulgo_date = Rs_chul("chulgo_date")
        chulgo_stock = Rs_chul("chulgo_stock")
        chulgo_stock_name = Rs_chul("chulgo_stock_name")
	    chulgo_stock_company = Rs_chul("chulgo_stock_company")
	    rele_att_file = Rs_chul("rele_att_file")
	    rele_memo = Rs_chul("rele_memo")
        rele_sign_yn = Rs_chul("rele_sign_yn")
	    rele_sign_no = Rs_chul("rele_sign_no")
	    rele_sign_date = Rs_chul("rele_sign_date")
	    if chulgo_date = "0000-00-00" then
	          chulgo_date = ""
	    end if
   else
		rele_stock = ""
        rele_seq = ""
	    rele_date = ""
        rele_id = ""
        rele_goods_type = ""
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
        chulgo_ing = ""
        chulgo_date = ""
        chulgo_stock = ""
        chulgo_stock_name = ""
	    chulgo_stock_company = ""
	    rele_att_file = ""
	    rele_memo = ""
        rele_sign_yn = ""
	    rele_sign_no = ""
	    rele_sign_date = ""
end if
Rs_chul.close()

sql = "select * from met_mv_reg_goods where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')  ORDER BY rl_goods_seq,rl_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "â���̵� ����Ƿ� ��ȸ"

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
					document.frm.action = "met_move_reg_approve_ok.asp";
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
				<form method="post" name="frm" action="met_move_reg_cancel.asp">
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
                                <th>��û����</th>
							    <td class="left"><%=rele_date%></td>
							    <th>�뵵����</th>
							    <td class="left"><%=rele_goods_type%></td>
							    <th>��ûâ��</th>
							    <td class="left"><%=rele_stock_name%>&nbsp;(<%=rele_stock_company%>)</td>
 							</tr>
                            <tr>
							    <th>ȸ��</th>
							    <td class="left"><%=rele_company%></td>
							    <th>�����</th>
							    <td class="left"><%=rele_saupbu%></td>
							    <th>��û��</th>
							    <td class="left"><%=rele_emp_name%>&nbsp;(<%=rele_org_name%>)</td>
						    </tr>
							<tr>
                                <th>����û��</th>
							    <td class="left"><%=chulgo_rele_date%></td>
							    <th>���óâ��</th>
							    <td colspan="3" class="left"><%=chulgo_stock_name%>&nbsp;(<%=chulgo_stock_company%>)</td>
						    </tr>
                            <tr>
                                <th>�������</th>
							    <td class="left"><%=chulgo_date%>&nbsp;</td>
							    <th>��ûâ��<br>�԰���</th>
							    <td colspan="3" class="left"><%=in_stock_date%>&nbsp;</td>
						    </tr>
							<tr>
							  <th>���</th>
							  <td colspan="5" class="left"><%=rele_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� â���̵� ����Ƿ� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="8%" >
                            <col width="*" >
                            <col width="12%" >
							<col width="18%" >
							<col width="18%" >
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
                                <td><%=rs("rl_goods_type")%>&nbsp;</td>
                                <td><%=rs("rl_goods_grade")%>&nbsp;</td>
								<td><%=rs("rl_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("rl_goods_code")%>&nbsp;</td>
                                <td><%=rs("rl_goods_name")%>&nbsp;</td>
                                <td><%=rs("rl_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("rl_qty"),0)%>&nbsp;</td>
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
                            <tr>
							  <th>÷��</th>
							  <td colspan="5" class="left">
                        <% 
                           If rele_att_file <> "" Then 
                              path = "/met_upload/" 
                        %>
                              <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rele_att_file%>"><%=rele_att_file%></a>
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
                        <% if rele_sign_yn = "N" then	%>
                            <span class="btnType01"><input type="button" value="�����û" onclick="javascript:approve_request('<%=rele_date%>','<%=rele_stock%>','<%=rele_seq%>');"></span>
                        <% end if	%>
                        <% if cancel_yn = "Y" then	%>
                            <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();"></span>
                        <% end if	%>
                            <span class="btnType01"><input type="button" value="���" onclick="javascript:printWindow();"></span>
                            <span class="btnType01"><input type="button" value="�ݱ�" onclick="javascript:goAction();"></span>
                            <span class="btnType01"><input type="button" value="�Ƿڼ� ���" onClick="pop_Window('met_move_referral_print.asp?rele_stock=<%=rele_stock%>&rele_date=<%=rele_date%>&rele_seq=<%=rele_seq%>','met_move_referral_print_pop','scrollbars=yes,width=750,height=600')"></span>
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
