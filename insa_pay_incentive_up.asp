<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

    dim abc,filenm
    dim month_tab(24,2) ' �ͼӳ�� �޺��ڽ��� �����ϴ� ����
    Set abc = Server.CreateObject("ABCUpload4.XForm")
    abc.AbsolutePath = True
    abc.Overwrite = true
    abc.MaxUploadSize = 1024*1024*50

    pay_company = abc("pay_company")
    pay_month   = abc("pay_month")
    give_date   = abc("give_date")
    file_type   = abc("file_type")

    if ck_sw = "y" then
        pay_company = request("pay_company")
        pay_month   = request("pay_month")
    end if

    if pay_company = "" then
        ck_sw = "y"
    else
        ck_sw = "n"
    end if
        
    if pay_company = "" then
        pay_company = "���̿��������"
        curr_dd     = cstr(datepart("d",now)) ' ���糯¥(��)
        give_date   = mid(cstr(now()),1,10)
        from_date   = mid(cstr(now()-curr_dd+1),1,10) ' ������ ù��(1��)
        pay_month   = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
    end if
        
    '  �ͼӳ�� �޺��ڽ��� ���� [����]
    cal_month  = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
    view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
    month_tab(24,1) = cal_month
    month_tab(24,2) = view_month

    for i = 1 to 23
        cal_month = cstr(int(cal_month) - 1)
        if mid(cal_month,5) = "00" then
            cal_year  = cstr(int(mid(cal_month,1,4)) - 1)
            cal_month = cal_year + "12"
        end if	 
        view_month = mid(cal_month,1,4) + "�� " + mid(cal_month,5,2) + "��"
        
        j = 24 - i
        month_tab(j,1) = cal_month
        month_tab(j,2) = view_month
    next
    '  �ͼӳ�� �޺��ڽ��� ���� [��]

	
	Set DbConn = Server.CreateObject("ADODB.Connection")
	set cn     = Server.CreateObject("ADODB.Connection")
	set rs      = Server.CreateObject("ADODB.Recordset")	
	Set Rs_etc  = Server.CreateObject("ADODB.Recordset")
	Set Rs_org  = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp  = Server.CreateObject("ADODB.Recordset")
	Set Rs_bnk  = Server.CreateObject("ADODB.Recordset")
	Set Rs_give = Server.CreateObject("ADODB.Recordset")
    Set Rs_dct  = Server.CreateObject("ADODB.Recordset")
    Set rs_com  = Server.CreateObject("ADODB.Recordset")
    
	DbConn.Open dbconnect
'Response.write ck_sw&"<br><br>"
	If ck_sw = "n" Then
		Set filenm = abc("att_file")(1)
		
		path = Server.MapPath ("/pay_file")
		filename  = filenm.safeFileName
		fileType  = mid(filename,inStrRev(filename,".")+1)
		file_name = pay_company + "_" + pay_month + "_�󿩱�" + give_date
		
		save_path = path & "\" & file_name&"."&fileType
		if fileType = "xls" or fileType = "xlk" then
			file_type = "Y"
			filenm.save save_path
		
            objFile = save_path
            
			cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
			rs.Open "select * from [1:10000]",cn,"0"
				
			rowcount=-1
            xgr = rs.getrows
			rowcount = ubound(xgr,2)
			fldcount = rs.fields.count
            tot_cnt = rowcount + 1
		else
			objFile = "none"
			rowcount=-1
			file_type = "N"
		end if		  
	else
		objFile = "none"
		rowcount=-1
	end if
	title_line = "�󿩱� �ڷ� ���ε�"


    etc_code = "9999"	
    sql = "select * from emp_etc_code where emp_etc_code = '" + etc_code + "'"
    'Response.write Sql&"<br>"
    Rs_etc.Open Sql, Dbconn, 1

        emp_payend_date = Rs_etc("emp_payend_date")
        emp_payend_yn   = Rs_etc("emp_payend_yn")

    Rs_etc.close()

    'Response.write pay_month & "<br>"
    'Response.write emp_payend_date & "<br>"

    if pay_month > emp_payend_date then
        emp_payend = "N"
    else   
        emp_payend = "Y"
    end if   	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
    <title>�޿����� �ý���</title>
    <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
    <link href="/include/style.css" type="text/css" rel="stylesheet">
    <script src="/java/jquery-1.9.1.js"></script>
    <script src="/java/jquery-ui.js"></script>
    <script src="/java/common.js" type="text/javascript"></script>
    <script src="/java/ui.js" type="text/javascript"></script>
    <script type="text/javascript" src="/java/js_form.js"></script>
    <script type="text/javascript" src="/java/js_window.js"></script>
    
    <script type="text/javascript">

        // �˻� ��ư Ŭ��!!
        function frmcheck () 
        {
            if (chkfrm()) {
                document.frm.submit ();
            }
        }
        
        function chkfrm() {
            if (document.frm.pay_company.value == "") {
                alert ("ȸ�縦 �����ϼ���");
                return false;
            }	
            if (document.frm.pay_month.value == "") {
                alert ("�ͼӳ���� �����ϼ���");
                return false;
            }	
            if (document.frm.att_file.value == "") {
                alert ("���ε� ���� ������ �����ϼ���");
                return false;
            }	
            return true;
        }

        // �󿩱� upload ��ư Ŭ��!!
        function frm1check () 
        {
            if (chkfrm1()) {
                document.frm1.submit ();
            }
        }
        
        function chkfrm1() 
        {
            if (confirm('DB�� ���ε� �Ͻðڽ��ϱ�?')==true) {
                return true;
            }
            return false;
        }
        
        // �󿩱� Upload ���� ��ư Ŭ��!!
        function pay_month_updel(val, val2) 
        {
            if (!confirm("�󿩱� Upload�ڷḦ ���� �Ͻðڽ��ϱ� ?")) return;

            var frm = document.frm;
            
            document.frm.pay_month1.value   = document.getElementById(val).value;
            document.frm.pay_company1.value = document.getElementById(val2).value;
            
            document.frm.action = "insa_pay_incentive_up_del.asp";
            document.frm.submit();
        }	
    </script>
</head>
<body>
	<div id="wrap">			
	    <!--#include virtual = "/include/insa_pay_header.asp" -->
        <!--#include virtual = "/include/insa_pay_menu.asp" -->
        
		<div id="container">
            <h3 class="insa"><%=title_line%> ..������..</h3>
            
            <form action="insa_pay_incentive_up.asp" method="post" name="frm" enctype="multipart/form-data">
                
                <fieldset class="srch">
                    <legend>��ȸ����</legend>
                    <dl>
                        <dt>���ε峻��</dt>
                        <dd>
                        <p>
                            <label>
                                <strong>ȸ��: </strong>
                                <%
                                Sql="select * from emp_org_mst where org_level = 'ȸ��' ORDER BY org_code ASC"
                                rs_org.Open Sql, Dbconn, 1	
                                %>
                                <select name="pay_company" id="pay_company" type="text" style="width:110px">
                                    <option value="">����</option>
                                    <% 
                                    do until rs_org.eof 
                                        %>
                                        <option value='<%=rs_org("org_name")%>' <%If pay_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                                        <%
                                        rs_org.movenext()  
                                    loop 
                                    rs_org.Close()
                                    %>
                                </select>
                            </label>
                            <label>
                                <strong>�ͼӳ��: </strong>
                                <select name="pay_month" id="pay_month" type="text" value="<%=pay_month%>" style="width:90px">
                                <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pay_month = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                <%	next	%>
                                </select>
                            </label>
                            <br>
                            <label>
                                <strong>���ε�����: </strong>
                                <input name="att_file" type="file" id="att_file" size="100" value="<%=att_file%>" style="text-align:left"> 
                            </label>
                            <input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">

                            <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                        </p>
                        </dd>
                    </dl>
                </fieldset>
                
                <div class="gView">
                    <table cellpadding="0" cellspacing="0" class="tableList">
                        <colgroup>
                            <col width="3%" >
                            <col width="3%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="8%" >
                            <col width="3%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="8%" >
                        </colgroup>
                        <thead>
                            <tr>
                                <th class="first" scope="col">�Ǽ�</th>
                                <th scope="col">���</th>
                                <th scope="col">���</th>
                                <th scope="col">����</th>
                                <th scope="col">���޾� ��</th>
                                <th scope="col">��뺸��</th>			
                                <th scope="col">�ҵ漼</th>
                                <th scope="col">����ҵ漼</th>
                                <th scope="col">���� �հ�</th>
                            </tr>
                        </thead>
                        <tbody>
                            <%
                            tot_emp = 0
                            tot_bank = 0
                            tot_err = 0

                            tot_give_total = 0
                                
                            if rowcount > -1 then
                                for i=2 to (rowcount-1)
                                    if xgr(1,i) = "" or isnull(xgr(1,i)) then
                                        exit for ' �ڷᰡ ���̻� ������ �������´�.
                                    end if

                                    ' ���üũ 				
                                    emp_sw = "Y"
                                    emp_no = xgr(2,i)
                                    Sql = "select * from emp_master where emp_no = '"&xgr(2,i)&"'"
                                    Set rs_emp = DbConn.Execute(Sql)
                                    'Response.write Sql & "<br>"
                                    if rs_emp.eof then
                                        tot_emp = tot_emp + 1
                                        tot_err = tot_err + 1
                                        emp_sw = "N"
                                        emp_name ="[�̵��]"
                                    else
                                        emp_name = rs_emp("emp_name")	  
                                    end if
                                    'Response.write emp_sw & "<br>"
                                    name_sw = "Y"

                                    ' �������üũ
                                    bank_sw = "Y"
                                    Sql = "SELECT * FROM pay_bank_account where emp_no = '"&emp_no&"'"
                                    Set rs_bnk = DbConn.Execute(SQL)
                                    if  rs_bnk.eof then
                                        tot_bank = tot_bank + 1
                                        tot_err = tot_err + 1
                                        bank_sw = "N"

                                        emp_name = emp_name & " [������¹̵��]"
                                    end if
                                    rs_bnk.close()	 

                                    ' �����׸�
                                    pmg_give_total     = toString(xgr(11,i),"0") ' ���޾װ�
                                    
                                    ' �����׸�			
                                    de_epi_amt		    = toString(xgr(12,i),"0")  ' ��뺸��
                                    de_income_tax	    = toString(xgr(13,i),"0")  ' �ҵ漼
                                    de_wetax		    = toString(xgr(14,i),"0")  ' ����ҵ漼                                        
                                    de_deduct_total     = toString(xgr(14,i),"0")  ' �����װ�

                                    sql = "select * from pay_month_give where pmg_yymm = '"&pay_month&"' and pmg_id = '2' and pmg_emp_no = '"&emp_no&"'"
                                    set Rs_give=dbconn.execute(sql)
                                    'Response.write sql&"<br>"
                                    if Rs_give.eof or Rs_give.bof then
                                        reg_sw = "N"
                                    else
                                        reg_sw = "Y"
                                    end if
                                    
                                    tot_give_total 	 = tot_give_total   + pmg_give_total                                        
                                    tot_epi_amt 	 = tot_epi_amt      + de_epi_amt
                                    tot_income_tax 	 = tot_income_tax   + de_income_tax
                                    tot_wetax 	     = tot_wetax        + de_wetax
                                    tot_deduct_total = tot_deduct_total + de_deduct_total
                                    
                                    if reg_sw = "N" then 
                                        reg_flag = "No"
                                        bgcolor0=""
                                    else
                                        reg_flag = "Yes"
                                        bgcolor0="#FFCCFF"
                                    end if
                                    
                                    if (emp_sw = "Y") and (bank_sw = "Y") then
                                        bgcolor1=""
                                    else
                                        bgcolor1="#FFCCFF"
                                    end if
                                    
                                    if name_sw = "Y" then
                                        bgcolor2=""
                                    else
                                        bgcolor2="#FFCCFF"
                                    end if
                                    
                                    %>
                                    <tr>
                                        <td class="first"><%=i-1%></td>
                                        <td bgcolor="<%=bgcolor1%>"><%=reg_flag%></td>                            
                                        <td bgcolor="<%=bgcolor1%>"><%=emp_no%></td>
                                        <td bgcolor="<%=bgcolor1%>"><%=emp_name%></td>
                                        <td bgcolor="<%=bgcolor1%>" class="right"><%=formatnumber(pmg_give_total,0)%></td>
                                        <td bgcolor="<%=bgcolor1%>" class="right"><%=formatnumber(de_epi_amt,0)%></td>
                                        <td bgcolor="<%=bgcolor1%>" class="right"><%=formatnumber(de_income_tax,0)%></td>
                                        <td bgcolor="<%=bgcolor1%>" class="right"><%=formatnumber(de_wetax,0)%></td>
                                        <td bgcolor="<%=bgcolor1%>" class="right"><%=formatnumber(de_deduct_total,0)%></td>
                                    </tr>
                                    <%
                                next
                            end if
                            %>
                            <tr>
                                <th class="first">����</th>
                                <th colspan="3" title="�Ǽ�">�޿����¹̵��:<%=formatnumber(tot_bank,0)%> �����̵��:<%=formatnumber(tot_emp,0)%></th>
                                <th class="right"><%=formatnumber(tot_give_total,0)%></th>
                                <th class="right"><%=formatnumber(tot_epi_amt,0)%></th>
                                <th class="right"><%=formatnumber(tot_income_tax,0)%></th>
                                <th class="right"><%=formatnumber(tot_wetax,0)%></th>
                                <th class="right"><%=formatnumber(tot_deduct_total,0)%></th>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="15%"><div class="btnCenter"></div></td>
                    <td>
                        <div class="btnRight"><a href="#" onClick="pay_month_updel('pay_month','pay_company');return false;" class="btnType04">�󿩱� Upload ����</a></div>
                    </td> 
                </tr>
                </table>

                <input type="hidden" name="pay_company1" value="<%=pay_company%>" ID="Hidden1">
                <input type="hidden" name="pay_month1" value="<%=pay_month%>" ID="Hidden1">             
            </form>
            
            <%
            if emp_payend = "N" then 
                if tot_cnt <> 0 and tot_err = 0 then 
                %>
                <form action="insa_pay_incentive_up_ok.asp" method="post" name="frm1">
                    <br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="�󿩱��ڷ� Upload" onclick="javascript:frm1check();"NAME="Button1"></span>
                    </div>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
                    <input name="pmg_yymm" type="hidden" id="pmg_yymm" value="<%=pay_month%>">
                    <input name="pmg_date" type="hidden" id="pmg_date" value="<%=give_date%>">
                    <input name="pmg_company" type="hidden" id="pmg_company" value="<%=pay_company%>">
                    <br>
                </form>
                <%
                end if
            end if 
            %>
        </div>				
    </div>
</body>
</html>

