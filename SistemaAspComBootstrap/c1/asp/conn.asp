<%

	Dim codigoSQL
	Dim codigoInserir
	Dim codigoConnection
	Dim connection
	Dim returnMysql
	Dim updateCodigo
	Dim deleteCodigo


	codigoSQL = "SELECT * FROM cliente"
	codigoInserir = "INSERT INTO `felipe`.`cliente` (`Usuario`, `Senha`) VALUES ('" & Request.Form("nome") & "', '" & Request.Form("senha") & "');"
	codigoConnection = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=localhost; DATABASE=felipe; UID=root; PASSWORD=; OPTION=3"
	updateCodigo = "UPDATE `felipe`.`cliente` SET `Usuario` = '" & Request.Form("nomeEdit") & "', `Senha` = '" & Request.Form("senhaEdit") & "' WHERE `cliente`.`ID` = " & Request.Form("idEdit") & ";"
	deleteCodigo = "DELETE FROM `felipe`.`cliente` WHERE `cliente`.`ID` = " & Request.Form("idExcluir")

	Set connection = Server.CreateObject("ADODB.Connection")
	Set returnMysql = Server.CreateObject("ADODB.Recordset")

	connection.Open codigoConnection
	connection.Execute(codigoInserir)
	returnMysql.Open codigoSQL, connection


	If IsNull(Request.Form("idEdit")) Then
		Response.Write("Teste")
		If returnMysql.EOF Then
			Response.Write "O banco de dados esta vasio"
		Else
			Do While NOT returnMysql.EOF 
				If (StrComp (returnMysql("ID"), Request.Form("idEdit"),0)) = 1 Then
					connection.Execute(updateCodigo)
				End If
				returnMysql.MoveNext
			Loop

		End If
	End If

	If IsNull(Request.Form("idExcluir")) Then
		If returnMysql.EOF Then
			Response.Write "O banco de dados esta vasio"
		Else

			Do While NOT returnMysql.EOF 
				If (StrComp (returnMysql("ID"), Request.Form("idExcluir"),0)) = 1 Then
					connection.Execute(deleteCodigo)
				End If 
				returnMysql.MoveNext
			Loop

		End If
	End If

	If returnMysql.EOF Then
		Response.Write "O banco de dados esta vasio"
	Else
		Response.Write "<table>"
		Response.Write "<caption>Tabela do banco de dados</caption>"
		Do While NOT returnMysql.EOF 
			'Response.Write returnMysql("ID") & "<br/>"
			'Response.Write returnMysql("Usuario") & "<br/>"
			'Response.Write returnMysql("Senha") & "<br/>"
			'Response.Write "<br/>"
			Response.Write "<tr><td>" & returnMysql("ID") & "</td><td>" & returnMysql("Usuario") & "</td><td>" & returnMysql("Senha") & "</td></tr>"
			returnMysql.MoveNext
		Loop
		Response.Write "</table>"
	End If

	Dim vUsuario
	Dim vSenha

	If returnMysql.EOF Then
		'Response.Write "O banco de dados esta vasio Teste"
	Else
		Do While NOT returnMysql.EOF 
			If (StrComp(Request.Form("usuarioEntrar"), returnMysql("Usuario"))) = 1 Then
				vUsuario = "true"
				returnMysql.MoveNext
				If (StrComp(Request.Form("senhaEntrar"), returnMysql("Senha"))) = 1 Then
					vSenha = "true"
					
				End If
			End If
			returnMysql.MoveNext
		Loop

	End If

	connection.Close
	Set connection = Nothing
	Set returnMysql = Nothing
	

%>
<!-- 
<script type="text/javascript">
	var vUsuario = <%= vUsuario %>;
	var vSenha = <%= vSenha %>;
	if(vUsuario == true && vSenha == true){
		window.location.replace("sejaBemVindo.asp");
	}

</script>
 -->
<a href="edit.asp">Clique aqui caso voce queira editar</a><br/>
<a href="exclude.asp">Clique aqui caso voce queira Excluir</a><br/>
<a href="form.asp">Clique aqui para cadastrar mais</a><br/>
