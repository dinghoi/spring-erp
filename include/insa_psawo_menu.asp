<div class="btnRight">
<%
'in_name = request.cookies("nkpmg_user")("coo_user_name")
'in_empno = request.cookies("nkpmg_user")("coo_user_id")
%>
    <a href="/person/insa_individual_confirm.asp" class="btnType01">������ �߱�</a>
    <a href="/person/insa_confirm_report.asp" class="btnType01">������ �߱���Ȳ</a>
<%If SysAdminYn = "Y" Then%>
    <a href="insa_sawo_join.asp" class="btnType01"><span style="color:red;">����ȸ ����</span></a>
    <a href="insa_sawo_ask.asp" class="btnType01"><span style="color:red;">������ ��û</span></a>
    <a href="#" onClick="pop_Window('insa_sawo_ask.asp?sawo_empno=<%=in_empno%>&emp_name=<%=in_name%>&u_type=<%=""%>','insa_sawo_ask_pop','scrollbars=yes,width=750,height=400')" class="btnType01"><span style="color:red;">�����ݽ�û</span></a>
    <a href="insa_sawo_ask_report.asp" class="btnType01"><span style="color:red;">�����ݽ�û��Ȳ</span></a>
     <a href="insa_name_card.asp" class="btnType01"><span style="color:red;">���� ��û</span></a>
<%End If%>
</div>
