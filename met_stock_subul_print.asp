<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

view_condi=Request("view_condi")
goods_type=request("goods_type")
owner_view=request("owner_view")
condi=request("condi")
stock=request("stock")


If view_condi = "" Then
	view_condi = "���̿��������"
	stock = ""
	goods_type = "��ǰ"
	owner_view = "C"
	ck_sw = "n"
	condi = ""
End If

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY stock_company,stock_goods_grade,stock_goods_gubun,stock_goods_name,stock_goods_standard,stock_code ASC"
if goods_type = "��ü" then
   if condi = "" then
         where_sql = " WHERE (stock_company = '"&view_condi&"')" 
      else  
      if owner_view = "C" then 
             where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
         else
		     where_sql = " WHERE (stock_company = '"&view_condi&"') and (stock_goods_code = '"+condi+"')"
	   end if
   end if   
  else
   if condi = "" then
         where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"')" 
      else  
      if owner_view = "C" then 
             where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_name like '%"+condi+"%')"
         else
		     where_sql = " WHERE (stock_goods_type = '"&goods_type&"') and (stock_company = '"&view_condi&"') and (stock_goods_code = '"+condi+"')"
	   end if
   end if  
end if 

'if stock = "" then
'       stock_sql = ""
'   else
'       stock_sql = " and (stock_code = '"&stock&"') "
'end if

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (stock_name like '%"&stock&"%') "
end if

sql = "select * from met_stock_gmaster " + where_sql + stock_sql + order_sql
Rs.Open Sql, Dbconn, 1
'response.write(sql)

title_line = " ǰ�� ������Ȳ "

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
				<form action="met_stock_subul_print.asp" method="post" name="frm">
				<div class="gView">
				<table width="1150" cellpadding="0" cellspacing="0">
				  <tr>
				    <td height="50px" align="center"><font style="font-size:32px"><strong><%=title_line%></strong></td>
			      </tr>
                  <tr>
				    <td height="20px" align="left" width="85%"><font style="font-size:12px"><%=view_condi%></td>
                    <td height="20px" align="right" width="15%" ><font style="font-size:12px"><%=mid(cstr(now()),1,4)%>��&nbsp;<%=mid(cstr(now()),6,2)%>��&nbsp;<%=mid(cstr(now()),9,2)%>�� ����</td>
			      </tr>
				  </table>
					<br>
                <table width="1150" border="1px" cellpadding="0" cellspacing="0" bordercolor="#000000" class="tablePrt">
						<colgroup>
                              <col width="7%" >
                              <col width="7%" >
                              <col width="12%" >
                              <col width="14%" >
                              <col width="3%" > 
                              <col width="6%" >
				              <col width="8%" >
                              <col width="6%" >
				              <col width="8%" >
                              <col width="6%" >
				              <col width="8%" >
                              <col width="6%" >
				              <col width="8%" >
                        </colgroup>
						 <thead>
                              <tr bgcolor="#f8f8f8">
				                <th rowspan="2" height="30" class="first" align="center" scope="col" class="style14C">�ڵ�</th>
				                <th rowspan="2" scope="col" align="center" class="style14C">ǰ�񱸺�</th>
                                <th rowspan="2" scope="col" align="center" class="style14C">ǰ���</th>
                                <th rowspan="2" scope="col" align="center" class="style14C">�԰�</th>
                                <th rowspan="2" scope="col" align="center" class="style14C">����</th>
                                <th colspan="2" scope="col" align="center" class="style14C">����</th>
				                <th colspan="2" scope="col" align="center" class="style14C">�԰�</th>
                                <th colspan="2" scope="col" align="center" class="style14C">���</th>
                                <th colspan="2" scope="col" align="center" class="style14C">�⸻</th>
			                  </tr>
                              <tr>
				                <th scope="col" height="30" align="center" class="style14C" style=" border-left:1px solid #e3e3e3;">����</th>
                                <th scope="col" align="center" class="style14C">�ݾ�</th>
                                <th scope="col" align="center" class="style14C">����</th>
                                <th scope="col" align="center" class="style14C">�ݾ�</th>
                                <th scope="col" align="center" class="style14C">����</th>
                                <th scope="col" align="center" class="style14C">�ݾ�</th>
                                <th scope="col" align="center" class="style14C">����</th>
                                <th scope="col" align="center" class="style14C">�ݾ�</th>
                              </tr>
                        </thead>
						<tbody>
				     <%
						do until rs.eof or rs.bof
                             
	           		 %>
                            <tr>
				               <td height="20" align="center"><%=rs("stock_goods_code")%>&nbsp;</td>
                               <td align="center"><%=rs("stock_goods_gubun")%>&nbsp;</td>
                               <td align="center"><%=rs("stock_goods_name")%>&nbsp;</td>
                               <td align="center"><%=rs("stock_goods_standard")%>&nbsp;</td>
                               <td align="center"><%=rs("stock_goods_grade")%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_last_qty"),0)%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_last_amt"),0)%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_in_amt"),0)%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_go_amt"),0)%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_jj_qty"),0)%>&nbsp;</td>
                               <td class="right"><%=formatnumber(rs("stock_jj_amt"),0)%>&nbsp;</td>
			                </tr>
					<%
							rs.movenext()
						loop
						rs.close()
					%>
				</div>
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
			</form>
		</div>				
	</div>        				
	</body>
</html>

