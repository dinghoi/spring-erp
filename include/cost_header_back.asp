			<div id="header">
				<h1><img src="/image/com_logo.jpg" alt="Ȩ������" width="116" height="30"/></h1>
				<h2><img src="/image/cost_title.gif" alt="��� ���� �ý���" width="189" height="25"/></h2>
				<div class="login">
					<p>
                    <strong><%=user_name%>&nbsp;<%=user_grade%>��</strong> �ȳ��ϼ���
					</p>
				</div>	
                <div id="gnb">
                    <ul>
                        <li class="dep1"><a href="person_cost_report.asp">����ϰ���</a></li>    
                        <li class="dep1"><a href="emp_cost_mg.asp">�����Ȳ����</a></li>                        
                        <li class="dep1">                       
                    <%	if cost_grade = "0" or (bonbu = "ITO �������" and position = "�������") or (bonbu = "ITO �������" and position = "������") then	%>						
                        <a href="saupbu_profit_loss_total.asp">������Ȳ</a>
					<%	  else	%>
                        <a>������Ȳ</a>
					<%	end if	%>
						</li>
                        <li class="dep1">                       
                    <%	if cost_grade = "0" and (user_id = "100539" or user_id = "100031" or user_id = "100041" or user_id = "900001" or user_id = "100359") then	%>						
                        <a href="cost_grade_mg.asp">����ڵ����</a>
					<%	  else	%>
                        <a>����ڵ����</a>
					<%	end if	%>
						</li>
                    </ul>
                </div>
			</div>
