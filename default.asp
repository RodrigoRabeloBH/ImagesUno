<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
' Este arquivo deve estar no diretório que contenham subdiretorios com as imagens a serem exibidas ´só imagens(jpg, gif).
' Você pode criar um arquivo texto chamado imagens.txt para cada pasta
' sendo que o titulo de cada linha irá ser exibido na pagina
' defina o tamanho da borda para um valor maior que zero se quiser bordar ao redor da imagem
tamanho_borda = "5"
cor_borda = "aqua"
%>
<html>
<head>
<title>Galeria de Imagens</title>

<!--define o estilo a ser aplicado na página	-->
<style type="text/css">
	body {
	font-family: verdana;
	text-align: center;
		}
</style>

</head>
<a name="top"></a>
<h2>Galeria de Imagens</h2>
<body>
<%
'define as constantes usadas pelo objeto FileSystemObject usadas  no projeto
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

'recebe o diretório da requisição como 'f'
diretorio = request.querystring("f")

if diretorio = "" then

	pastaEspecifica = server.mappath(".")

	Set arquivoSistema = CreateObject("Scripting.FileSystemObject") 
	Set exemplo = arquivoSistema.GetFolder(pastaEspecifica) 
	Set colecaoPastas = exemplo.SubFolders

	For Each subPastas in colecaoPastas
		tamanhoPasta = left((subPastas.size/1000000), 3)
	    listaPasta = listaPasta & "<a href='?f=" & subPastas.name & "'><strong title='view'>&#187;</strong> " & subPastas.Name & " </a><small>&nbsp;(" & tamanhoPasta & " MB)</small>" & vbcrlf
	    listaPasta = listaPasta & "<BR>"  
	Next

	set arquivoSistema = nothing
	Response.Write listaPasta

else

caminhoArquivo = server.mappath(".") & "\" & diretorio
tituloArquivo = caminhoArquivo & "\imagens.txt"

Set arquivoSistema = CreateObject("Scripting.FileSystemObject")

Dim Vetor()

	If arquivoSistema.FileExists(tituloArquivo) then
		set file = arquivoSistema.GetFile(tituloArquivo)
		Set TextStream = file.OpenAsTextStream(ForReading,TristateUseDefault)
		contaTitulo = 0

		Do While Not TextStream.AtEndOfStream
			Linha = TextStream.readline
			ReDim Preserve Vetor(contaTitulo)
			Vetor(contaTitulo) = Linha
			'response.write contaTitulo & " " & Vetor(contaTitulo) & "<br>"
			contaTitulo = contaTitulo + 1
			'Response.write Linha
		Loop

		textStream.close

	end if

	Set exemplo = arquivoSistema.GetFolder(caminhoArquivo) 
	Set colecaoArquivos = exemplo.Files
	contaArquivo = 0
	
	For Each file in colecaoArquivos
	
		Ext = UCase(Right(File.Path, 3)) 
	
		If Ext = "JPG" OR Ext = "GIF" Then
	        on error resume next
	    	dados = Vetor(contaArquivo)
		    on error goto 0
		    caminhoReferencia = diretorio & "/" & file.name
		    caminhoImagem = "<strong>" & dados & "</strong><br><a href='" & caminhoReferencia & "' title='Galeria de Imagens' border=0><img src='" & caminhoReferencia & "' border='" & tamanho_borda & "' title=""" & dados & """ style='border-color: " & cor_borda & ";'></a><br>"
	        encheLista = encheLista & caminhoImagem & vbcrlf
	        encheLista = encheLista & "<BR>"
		    contaArquivo = contaArquivo + 1
		    dados = ""
		end if
	Next
	set arquivoSistema = Nothing
	encheLista = encheLista & "<br><small><a href='http://www.macoratti.net/indasp.htm' target='_blank'>Macoratti.net - Artigo - Galeria de Imagens</a></small>"
%>

<h3><a href="." title="up one level">&#171;</a>
&nbsp;<%=diretorio%></h3>

<p><%=encheLista%></p>

<% end if %>
<p style="font-size: xx-small;"><a href="#top" title="retorna ao topo">topo da página</a></p>
</body>
</html>

