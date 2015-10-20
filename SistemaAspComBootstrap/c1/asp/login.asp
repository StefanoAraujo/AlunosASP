<%
	
	Dim codigoConnection
	Dim returnMysql
	Dim codigoSQL
	
	codigoSQL = "SELECT * FROM usuario"
	codigoConnection = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=localhost; DATABASE=cadastrousuario; UID=root; PASSWORD=; OPTION=3"

	Set connection = Server.CreateObject("ADODB.Connection")
	Set returnMysql = Server.CreateObject("ADODB.Recordset")

	connection.Open codigoConnection
	returnMysql.Open codigoSQL, connection

	Dim vUsuario
	Dim vSenha

	If returnMysql.EOF Then
		'Response.Write "O banco de dados esta vasio Teste"
	Else
		Do While NOT returnMysql.EOF
			If StrComp(Request.Form("usuarioEntrar"), returnMysql("Usuario")) = 0 Then
				vUsuario = "true"
				If StrComp(Request.Form("senhaEntrar"), returnMysql("Senha")) = 0 Then
					vSenha = "true"
				End If
			End If
			returnMysql.MoveNext
		Loop

	End If

	If (StrComp(vUsuario, "true") = 0 and StrComp(vSenha, "true") = 0) Then
	Response.Write vUsuario
	Response.Write vSenha
		%>
			<script>
				window.location.replace("../dashboard.html");
			</script>
		<%
	Else
		Response.Write "Não foi posivel entrar no sistema redigite a senha" 
		%> <a href="../loginEntrar.html">Clique aqui para voltar</a> <%
	End If

	connection.Close
	Set connection = Nothing
	Set returnMysql = Nothing

%>
