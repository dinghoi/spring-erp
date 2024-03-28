<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim au_code_last

u_type  = request("u_type")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs     = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")					

DbConn.Open dbconnect

sql = "  SELECT *               " & chr(13) & _
      "    FROM as_unitprice    " & chr(13) & _
      "   WHERE delete_yn = 'N' " & chr(13) & _
      "ORDER BY au_code ASC     "
Rs.Open Sql, Dbconn, 1

if u_type = "U" then
    au_code       = request("au_code")

    sql = "SELECT *                            " & chr(13) & _
          "  FROM as_unitprice                 " & chr(13) & _
          " WHERE au_code = '" & au_code & "'  " & chr(13) & _
          "   AND delete_yn = 'N'              "
    Set rs_etc=DbConn.Execute(Sql)
    
    au_code       = rs_etc("au_code")
    au_name       = rs_etc("au_name")
    cost_center   = rs_etc("cost_center")
    as_unitprice1 = rs_etc("as_unitprice1")
    as_unitprice2 = rs_etc("as_unitprice2")
else
    au_code       = ""
    au_name       = ""
    cost_center   = ""
    as_unitprice1 = 0
    as_unitprice2 = 0
end if	

title_line = "AS ǥ�شܰ� ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>��� ���� �ý���</title>
        
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

        <script type="text/javascript">
        
            $(document).ready(function(){
				<%
                message = Request("message")

                if  (message <> "") then 
                    %>alert('<%=message%>');<%
                end if
                %>
			});

            function ApplyTo() 
            {
                document.apply_frm.submit();
            }

			function frmcheck () {
				if (chkfrm() == true) {
					document.frm.submit();
				}
			}
			
            function chkfrm() 
            {
                message = "";

                if (document.frm.au_code.value == "")
                {
                    message += "�����ڵ带 �Է��Ͻñ� �ٶ��ϴ�\n";
                }
                else
                {
                    if (document.frm.au_code.value.length == 4) 
                    {
                        if  (document.frm.au_code.value.substr(0,2) != "AU")
                        {
                            message += "�����ڵ�� 'AU'�� �����Ͽ��� �մϴ�.\n";            
                        }
                    }
                    else
                    {
                        message += "�����ڵ�� 4�ڸ��̾�� �մϴ�.\n";       
                    }
                }
                if (document.frm.au_name.value == "")
                {
                    message += "������ �Է��Ͻñ� �ٶ��ϴ�\n";
                }
                if (eval(document.frm.as_unitprice1.value) <= 0)
                {
                    message += "ǥ�شܰ��� 0���� Ŀ�� �մϴ�.\n";
                }
                
                if (message != "") 
                {
					alert (message);
					return false;
                }
                else		
				    return true ;
            }

            function frmdelete() 
            {
                if ( confirm('���� �����Ͻðڽ��ϱ�?') == true) 
                {
                    document.frm.u_type.value = "D";
                    document.frm.submit();
                }
            }
            
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
                
                <form action="as_unitprice_apply.asp" method="post" name="apply_frm">
                <fieldset class="srch">
                    <legend>��ȸ����</legend>
                    <dl>					
                        <dt></dt>
                        <dd>
                            <p>
                                <strong>������ : </strong>
                                <select name="apply_year" id="apply_year" style="width:100px">
                                    <% 
                                    nYear = mid(cstr(now()),1,4) 

                                    for i = -1 to 1

                                        iYear = mid(cstr(DateAdd("yyyy",i,now())),1,4) 
                                        %>
                                        <option value="<%=iYear%>" <%If nYear = iYear then %>selected<% end if %>><%=iYear%></option>
                                        <%
                                    next
                                    %>    
                                </select>
                                <select name="apply_month" id="apply_month" style="width:50px">
                                    <% 
                                    nMonth = mid(cstr(now()),6,2) 

                                    for i = 1 to 12

                                        if len(cstr(i)) = 1 then
                                            iMonth = "0"&cstr(i)
                                        else
                                            iMonth = cstr(i)
                                        end if
                                        %>
                                        <option value="<%=iMonth%>" <%If nMonth = iMonth then %>selected<% end if %>><%=iMonth%></option>
                                        <%
                                    next
                                    %>    
                                </select>
                                    
                                <input type="button" value="��������" onclick="javascript:ApplyTo();">
                            </p>
                        </dd>
                    </dl>
                </fieldset>
                </form>

				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="64%" height="356" valign="top"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="9%" >
				          <col width="*" >
				          <col width="12%" >
				          <col width="12%" >
				          <col width="12%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">�����ڵ�</th>
				            <th scope="col">����</th>
				            <th scope="col">���ͼ�</th>
				            <th scope="col">ǥ�شܰ�</th>
				            <th scope="col">Ư���ܰ�</th>
			              </tr>
			            </thead>
			            <tbody>
                        <%                        
                        do until rs.eof
                            %>
                            <tr>
                                <td class="first"><%=rs("au_code")%></td>
                                <td><a href="as_unitprice_mg.asp?au_code=<%=rs("au_code")%>&u_type=<%="U"%>"><%=rs("au_name")%></a></td>
                                <td><%=rs("cost_center")%></td>
                                <td><%=formatnumber(rs("as_unitprice1"),0)%></td>
                                <td><%=formatnumber(rs("as_unitprice2"),0)%></td>
                            </tr>
                            <%							
						    rs.movenext()
						loop

						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="34%" valign="top">
                        <form method="post" name="frm" action="as_unitprice_save.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
                                <th width="25%">�����ڵ�</th>
                                <td class="left">
                                <%
                                if u_type = "U" then
                                    %><%=au_code%> <input name="au_code" type="hidden" value="<%=au_code%>"><%
                                else
                                    %><input name="au_code" type="text" value="" size="8" maxlength="4" ><%
                                end if
                                %>
                                </td>
                            </tr>
				            <tr>
                                <th>����</th>
                                <td class="left"><input name="au_name" type="text" id="au_name" value="<%=au_name%>" notnull errname="�ڵ��"></td>
			                </tr>
				            <tr>
                                <th>���ͼ�</th>
                                <td class="left">
                                    <select name="cost_center" id="cost_center" style="width:100px">
                                        <option value="�ι������" <% if cost_center = "�ι������" then %>selected<% end if %>>�ι������</option>
                                        <option value="��������" <% if cost_center = "��������" then %>selected<% end if %>>��������</option>
                                        <option value="����������" <% if cost_center = "����������" then %>selected<% end if %>>����������</option>
                                        <option value="������" <% if cost_center = "������" then %>selected<% end if %>>������</option>
                                    </select>
                                </td>
			                </tr>
				            <tr>
                                <th>ǥ�شܰ�</th>
                                <td class="left"><input name="as_unitprice1" type="text" id="as_unitprice1" value="<%=formatnumber(as_unitprice1,0)%>" onKeyUp="plusComma(this);" style="width:80px;text-align:right"></td>
                            </tr>
                            <tr>
                                <th>Ư���ܰ�</th>
                                <td class="left"><input name="as_unitprice2" type="text" id="as_unitprice2" value="<%=formatnumber(as_unitprice2,0)%>" onKeyUp="plusComma(this);" style="width:80px;text-align:right"></td>
                            </tr>
                            </tbody>
			            </table>
						<br>
                        <input type="hidden" name="u_type" value="<%=u_type%>">
                        
				        <div align=center>
                            <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();"></span>
                            <span class="btnType01"><input type="button" value="����" onclick="javascript:frmdelete();"></span>
                        </div>
                      </form>
                      </td>
			        </tr>
				    <tr>
				      <td width="49%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="49%">&nbsp;</td>
			        </tr>
			      </table>
                </div>
			</div>				
	</div>        				
	</body>
</html>

