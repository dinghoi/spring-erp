<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

mvin_in_date = request("mvin_in_date")
mvin_in_stock = request("mvin_in_stock")
mvin_in_seq = request("mvin_in_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_stin = Server.CreateObject("ADODB.Recordset")
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_mv_in where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"')"
Set Rs_stin = DbConn.Execute(SQL)
if not Rs_stin.eof then
	
        mvin_id = Rs_stin("mvin_id")
        mvin_goods_type = Rs_stin("mvin_goods_type")
		mvin_stock_company = Rs_stin("mvin_stock_company")
		mvin_stock_name = Rs_stin("mvin_stock_name")
        mvin_emp_no = Rs_stin("mvin_emp_no")
        mvin_emp_name = Rs_stin("mvin_emp_name")
        mvin_company = Rs_stin("mvin_company")
        mvin_bonbu = Rs_stin("mvin_bonbu")
        mvin_saupbu = Rs_stin("mvin_saupbu")
        mvin_team = Rs_stin("mvin_team")
        mvin_org_name = Rs_stin("mvin_org_name")
        rele_date = Rs_stin("rele_date")
		rele_stock = Rs_stin("rele_stock")
		rele_seq = Rs_stin("rele_seq")
		chulgo_date = Rs_stin("chulgo_date")
		chulgo_stock = Rs_stin("chulgo_stock")
		chulgo_seq = Rs_stin("chulgo_seq")
		
		mvin_no = mid(cstr(Rs_stin("mvin_in_date")),3,2) + mid(cstr(Rs_stin("mvin_in_date")),6,2) + mid(cstr(Rs_stin("mvin_in_date")),9,2) 
		chulgo_no = mid(cstr(Rs_stin("chulgo_date")),3,2) + mid(cstr(Rs_stin("chulgo_date")),6,2) + mid(cstr(Rs_stin("chulgo_date")),9,2) 
		rele_no = mid(cstr(Rs_stin("rele_date")),3,2) + mid(cstr(Rs_stin("rele_date")),6,2) + mid(cstr(Rs_stin("rele_date")),9,2) 
   else
		mvin_id = ""
        mvin_goods_type = ""
		mvin_stock_company = ""
		mvin_stock_name = ""
        mvin_emp_no = ""
        mvin_emp_name = ""
        mvin_company = ""
        mvin_bonbu = ""
        mvin_saupbu = ""
        mvin_team = ""
        mvin_org_name = ""
        rele_date = ""
		rele_stock = ""
		rele_seq = ""
		chulgo_date = ""
		chulgo_stock = ""
		chulgo_seq = ""
end if
Rs_stin.close()

sql = "select * from met_mv_go where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')"
Set Rs_chul = DbConn.Execute(SQL)
if not Rs_chul.eof then
	
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
		chulgo_memo = Rs_chul("chulgo_memo")
   else
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

sql = "select * from met_mv_in_goods where (mvin_in_date = '"&mvin_in_date&"') and (mvin_in_stock = '"&mvin_in_stock&"') and (mvin_in_seq = '"&mvin_in_seq&"')  ORDER BY in_goods_seq,in_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "â���̵� �԰� ��ȸ"

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
					document.frm.action = "met_move_stin_approve_ok.asp";
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
				<form method="post" name="frm" action="met_move_stin_cancel.asp">
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
							  <th>�԰���</th>
							  <td class="left"><%=mvin_in_date%></td>
                              <th>�԰��ȣ</th>
							  <td class="left"><%=mvin_no%>&nbsp;<%=mvin_in_stock%><%=mvin_in_seq%>&nbsp;</td>
							  <th>�԰���</th>
							  <td class="left"><%=mvin_emp_name%>(<%=mvin_emp_no%>)</td>
						    </tr>
                      <th>�Ƿ�����<br>���</th>
							  <td class="left"><%=rele_date%>&nbsp;<%=rele_emp_name%>(<%=rele_emp_no%>)</td>
                              <th>�뵵����</th>
							  <td class="left"><%=mvin_goods_type%>&nbsp;
                              <th>��ûâ��</th>
							  <td class="left"><%=rele_stock_name%>&nbsp;(<%=rele_stock_company%>)&nbsp;</td>
						    </tr>
                            <tr>
                              <th>�������</th>
						      <td class="left"><%=chulgo_date%>&nbsp;&nbsp;(��û:&nbsp;<%=chulgo_rele_date%>)</td>
							  <th>�����</th>
							  <td class="left"><%=chulgo_emp_name%>&nbsp;(<%=chulgo_org_name%>)</td>
							  <th>���â��</th>
							  <td class="left"><%=chulgo_stock_name%>&nbsp;(<%=chulgo_stock_company%>)</td>
 							</tr>
							<tr>
							  <th>���</th>
							  <td colspan="5" class="left"><%=chulgo_memo%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� â���̵� �԰� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
                            <col width="8%" >
                            <col width="*" >
                            <col width="15%" >
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
								<th scope="col">�԰����</th>
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
                                <td><%=rs("in_goods_type")%>&nbsp;</td>
                                <td><%=rs("in_goods_grade")%>&nbsp;</td>
								<td><%=rs("in_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("in_goods_code")%>&nbsp;</td>
                                <td><%=rs("in_goods_name")%>&nbsp;</td>
                                <td><%=rs("in_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("in_qty"),0)%>&nbsp;</td>
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
       <% if cancel_yn = "Y" then   %>
                            <span class="btnType01"><input type="button" value="�μ��� ���" onClick="pop_Window('met_move_stin_receip_print.asp?mvin_in_date=<%=mvin_in_date%>&mvin_in_stock=<%=mvin_in_stock%>&mvin_in_seq=<%=mvin_in_seq%>','met_move_stin_print_pop','scrollbars=yes,width=750,height=600')"></span>
       <% end if %>                     
                        </div>
					</div>
					<br>               		
                    <input type="hidden" name="mvin_in_date" value="<%=mvin_in_date%>">
					<input type="hidden" name="mvin_in_stock" value="<%=mvin_in_stock%>">
					<input type="hidden" name="mvin_in_seq" value="<%=mvin_in_seq%>">
					<input type="hidden" name="cancel_yn" value="<%=cancel_yn%>">      				
	     </form>
    	</div>				
	  </div>     
	</body>
</html>
