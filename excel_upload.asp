<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/srvmg_dbcon_db.asp" -->
<% 
'  xlspath = dir_root & "\export\xls\" & dbname & ".xls"
'	Response.Buffer = True
	tot_cnt = 0
	tot_err = 0
	tot_com = 0
	tot_dept = 0
	tot_cust = 0
	tot_ddd = 0
	tot_tel = 0
	tot_sido = 0
	tot_gugun = 0
	tot_dong = 0
	tot_addr = 0
	tot_ce = 0
	
	ck_sw=Request("ck_sw")
	
	If ck_sw = "n" Then
		dim abc,filenm
	
		Set abc = Server.CreateObject("ABCUpload4.XForm")
	
		abc.AbsolutePath = True
		abc.Overwrite = true
	
		Set filenm = abc("att_file")(1)
		
		path = Server.MapPath ("/srv_upload")
		filename = filenm.safeFileName
		
		save_path = path & "\" & filename
			
	' 	if filename = "" then 
'		if filenm.length < 4194304  then 
'			If filename <> "" Then
'				if filenm <> "" then filenm.save save_path
		filenm.save save_path


		objFile = save_path
'		objFile = Request.form("att_file")
'		objFile = SERVER.MapPath("att_file")
'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
'		response.write(objFile)
		set cn = Server.CreateObject("ADODB.Connection")
		set rs = Server.CreateObject("ADODB.Recordset")
	
		Set DbConn = Server.CreateObject("ADODB.Connection")
		Set Rs_etc = Server.CreateObject("ADODB.Recordset")
		DbConn.Open dbconnect
	
		cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
		rs.Open "select * from [1:10000]",cn,"0"
		
		rowcount=-1
		xgr = rs.getrows
		rowcount = ubound(xgr,2)
		fldcount = rs.fields.count
		tot_cnt = rowcount + 1
	  Else
		objFile = "none"
		rowcount=-1
	End if

%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="javascript" src="/java/PopupCalendar.js"></script>
<style type="text/css">
<!--
.style3 {font-size: 11px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; }
.style4 {font-size: 11px; font-family: "굴림체", "돋움체", Seoul; }
.style5 {font-size: 12px; font-weight: bold; font-family: "굴림체", "돋움체", Seoul; }
-->
</style>
</head>
<body>
<table width="800" border="0">
  <tr> 
    <td width="800" height="41"><img src="image/k1_excel_upload_title.gif" width="800" height="40"></td>
  </tr>
  <tr> 
    <td width="800" height="6"><form action="excel_upload.asp?ck_sw=n" method="post" name="form1" enctype="multipart/form-data">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="800" height="32"  border="1" cellpadding="0" cellspacing="0">
            <tr>
              <td width="180" height="30" align="center" valign="middle" class="style5"><div align="center">1. 업로드 EXCEL 파일 선택 </div></td>
              <td height="30" valign="middle"><input name="att_file" type="file" id="att_file" size="60" value="<%=objFile%>"></td>
              <td width="70" height="30"><div align="center">
                  <input name="imageField" type="image" src="image/burton/upbtn.gif" width="55" height="20" border="0">
              </div></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td height="1">&nbsp;</td>
        </tr>
        <tr>
          <td><table width="100%"  border="1" cellpadding="0" cellspacing="0">
            <tr bgcolor="#CCCCCC">
              <td width="30" height="25"><div align="center" class="style3">SEQ</div></td>
              <td width="90" height="25"><div align="center" class="style3">회사</div></td>
              <td width="100" height="25"><div align="center" class="style3">부서</div></td>
              <td width="50" height="25"><div align="center" class="style3">고객명</div></td>
              <td width="30" height="25"><div align="center" class="style3">DDD</div></td>
              <td width="35" height="25"><div align="center" class="style3">TEL</div></td>
              <td width="40" height="25"><div align="center" class="style3">TEL_NO</div></td>
              <td width="40" height="25"><div align="center" class="style3">시도</div></td>
              <td width="80" height="25"><div align="center" class="style3">구군</div></td>
              <td width="70" height="25"><div align="center" class="style3">동/읍</div></td>
              <td width="155" height="25"><div align="center" class="style3">번지</div></td>
              <td width="50" height="25" class="style3"><div align="center">CE</div></td>
            </tr>
            <%
	  if rowcount > -1 then
		for i=0 to rowcount
' 회사명
		com_sw = "Y"
		sql_etc = "select * from etc_code where etc_type = '51' and etc_name = '" + xgr(0,i) +"'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_com = tot_com + 1
			tot_err = tot_err + 1
			com_sw = "N"
		end if
' 부서명
		dept_sw = "Y"
		if xgr(1,i) = "" then
			dept_sw = "N"
			tot_dept = tot_dept + 1
			tot_err = tot_err + 1
		end if
' 고객명
		cust_sw = "Y"
		if xgr(2,i) = "" then
			cust_sw = "N"
			tot_cust = tot_cust + 1
			tot_err = tot_err + 1
		end if
' DDD
		ddd_sw = "Y"
		sql_etc = "select * from etc_code where etc_type = '71' and etc_name = '" + xgr(3,i) +"'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_ddd = tot_ddd + 1
			tot_err = tot_err + 1
			ddd_sw = "N"
		end if
' TEL (전화국)
		tel_sw = "Y"
		if xgr(4,i) < "100" then
			tel_sw = "N"
			tot_tel = tot_tel + 1
			tot_err = tot_err + 1
		end if
' TEL NO
		tel_no_sw = "Y"
		if xgr(5,i) < "0001" then
			tel_no_sw = "N"
			tot_tel = tot_tel + 1
			tot_err = tot_err + 1
		end if
' 시도
		sido_sw = "Y"
		sql_etc = "select * from etc_code where etc_type = '81' and etc_name = '" + xgr(6,i) +"'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_sido = tot_sido + 1
			tot_err = tot_err + 1
			sido_sw = "N"
		end if
' 구군
		gugun_sw = "Y"
		sql_etc = "select * from ce_area where sido = '" + xgr(6,i) +"' and gugun = '" + xgr(7,i) + "'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_gugun = tot_gugun + 1
			tot_err = tot_err + 1
			gugun_sw = "N"
			mg_ce_id = ""
		  else
			mg_ce_id = rs_etc("mg_ce_id")	  
		end if
' CE
		ce_sw = "Y"
		sql_etc = "select * from memb where user_id = '" + mg_ce_id + "'"
		set rs_etc=dbconn.execute(sql_etc)				
		if rs_etc.eof then
			tot_ce = tot_ce + 1
			tot_err = tot_err + 1
			ce_sw = "N"
			mg_ce = "미등록"
		  else
			mg_ce = rs_etc("user_name")
		end if
' 동/읍
		dong_sw = "Y"
		if xgr(8,i) = "" then
			dong_sw = "N"
			tot_dong = tot_dong + 1
			tot_err = tot_err + 1
		end if
' 번지
		addr_sw = "Y"
		if xgr(9,i) = "" then
			addr_sw = "N"
			tot_addr = tot_addr + 1
			tot_err = tot_err + 1
		end if

    %>
            <tr valign="middle">
              <td width="30" height="20"><div align="center" class="style4"><%=i+1%></div></td>
              <% if com_sw = "Y" then %>
              <td width="90" height="20"><div align="center" class="style4"><%=xgr(0,i)%></div></td>
              <% else %>
              <td width="90" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(0,i)%></div></td>
              <% end if %>
              <% if dept_sw = "Y" then %>
              <td width="100" height="20"><div align="center" class="style4"><%=xgr(1,i)%></div></td>
              <% else %>
              <td width="100" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(1,i)%></div></td>
              <% end if %>
              <% if cust_sw = "Y" then %>
              <td width="50" height="20"><div align="center" class="style4"><%=xgr(2,i)%></div></td>
              <% else %>
              <td width="50" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(2,i)%></div></td>
              <% end if %>
              <% if ddd_sw = "Y" then %>
              <td width="30" height="20"><div align="center" class="style4"><%=xgr(3,i)%></div></td>
              <% else %>
              <td width="30" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(3,i)%></div></td>
              <% end if %>
              <% if tel_sw = "Y" then %>
              <td width="35" height="20"><div align="center" class="style4"><%=xgr(4,i)%></div></td>
              <% else %>
              <td width="35" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(4,i)%></div></td>
              <% end if %>
              <% if tel_no_sw = "Y" then %>
              <td width="40" height="20"><div align="center" class="style4"><%=xgr(5,i)%></div></td>
              <% else %>
              <td width="40" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(5,i)%></div></td>
              <% end if %>
              <% if sido_sw = "Y" then %>
              <td width="40" height="20"><div align="center" class="style4"><%=xgr(6,i)%></div></td>
              <% else %>
              <td width="40" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(6,i)%></div></td>
              <% end if %>
              <% if gugun_sw = "Y" then %>
              <td width="80" height="20"><div align="center" class="style4"><%=xgr(7,i)%></div></td>
              <% else %>
              <td width="80" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(7,i)%></div></td>
              <% end if %>
              <% if dong_sw = "Y" then %>
              <td width="70" height="20"><div align="center" class="style4"><%=xgr(8,i)%></div></td>
              <% else %>
              <td width="70" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(8,i)%></div></td>
              <% end if %>
              <% if addr_sw = "Y" then %>
              <td width="155" height="20"><div align="center" class="style4"><%=xgr(9,i)%></div></td>
              <% else %>
              <td width="155" height="20" bgcolor="#FFCCFF"><div align="center" class="style4"><%=xgr(9,i)%></div></td>
              <% end if %>
              <% if ce_sw = "Y" then %>
              <td width="50" height="20" class="style4"><div align="center"><%=mg_ce%></div></td>
              <% else %>
              <td width="50" height="20" class="style4" bgcolor="#FFCCFF"><div align="center"><%=mg_ce%></div></td>
              <% end if %>
            </tr>
            <% 
		next
	  end if
	%>
          </table></td>
        </tr>
      </table>
      </form>	</td>
  </tr>
  <tr>
    <td height="2"><table width="800" border="1">
      <tr bgcolor="#CCFFFF" class="style3">
        <td width="60" height="20"><div align="center">총건수</div></td>
        <td width="64" height="20"><div align="center">총Error</div></td>
        <td width="60" height="20"><div align="center">회사</div></td>
        <td width="60" height="20"><div align="center">부서</div></td>
        <td width="60" height="20"><div align="center">고객명</div></td>
        <td width="60" height="20"><div align="center">DDD</div></td>
        <td width="60" height="20"><div align="center">TEL</div></td>
        <td width="60" height="20"><div align="center">시도</div></td>
        <td width="60" height="20"><div align="center">구군</div></td>
        <td width="60" height="20"><div align="center">동/읍</div></td>
        <td width="60" height="20"><div align="center">번지</div></td>
        <td width="60" height="20"><div align="center">CE</div></td>
      </tr>
      <tr class="style4">
        <td width="60" height="20"><div align="center"><%=tot_cnt%></div></td>
        <td width="64" height="20"><div align="center"><%=tot_err%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_com%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_dept%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_cust%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_ddd%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_tel%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_sido%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_gugun%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_dong%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_addr%></div></td>
        <td width="60" height="20"><div align="center"><%=tot_ce%></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="3"><div align="center">
		<% if tot_err = 0 then %>
      		<a href="excel_upload_ok.asp">DB저장</a>
		<% end if %>
	</div></td>
  </tr>
</table>
</body>
</html>
<%
'	response.flush
	if rowcount <> -1 then
		rs.close
		cn.close
		set rs = nothing
		set cn = nothing
	end if
%>