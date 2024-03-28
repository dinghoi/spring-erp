<%@LANGUAGE="VBSCRIPT"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<%
	pg_name = request("pg_name")
'	ck_sw = Request("ck_sw")
	from_date = request.form("from_date")
	to_date = request.form("to_date")
	as_type = request.form("as_type")
	company = request.form("company")
	reside = request.form("reside")
	belong = request.form("belong")
	team = request.form("team")
	acpt_place = request.form("acpt_place")
	mg_group = request.form("mg_group")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title>무제 문서</title>
<script>
function Submit_frm() 
{
	document.frm.submit ();
}
</script>
</head>

<style>html, body { width: 100%; height: 100%; } </style>

<body onLoad="Submit_frm();">
<table width="100%" height="100%"  border="0">
  <tr>
    <td align="center" valign="middle">
		<table width="222" BORDER=0 CELLPADDING=0 CELLSPACING=0>
		<tr>
        	<td>
                <div align="center"><img src="image/wait.gif" width="222" height="118">
                    <form name="frm" method="post" action="<%=pg_name%>">
                      <p>
                        <input name="from_date" type="hidden" id="from_date3" value="<%=from_date%>">
                        <input name="to_date" type="hidden" id="to_date" value="<%=to_date%>">
                        <input name="as_type" type="hidden" id="as_type2" value="<%=as_type%>">
                        <input name="company" type="hidden" id="company2" value="<%=company%>">
                        <input name="reside" type="hidden" id="reside" value="<%=reside%>">
                        <input name="belong" type="hidden" id="belong" value="<%=belong%>">
                        <input name="team" type="hidden" id="team" value="<%=team%>">
                        <input name="acpt_place" type="hidden" id="acpt_place" value="<%=acpt_place%>">
                        <input name="mg_group" type="hidden" id="mg_group" value="<%=mg_group%>">
                      </p>
                    </form>
                </div>
			</td>
        </tr> 
		</table>
    </td>
  </tr>
</table>
</body>
</html>
