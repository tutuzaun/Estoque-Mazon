' =============================================================================
' RESTAURANTE MAZON - Inicializador Silencioso (Sem Janela de Prompt / CMD)
' Executa o script iniciar_estoque.bat em segundo plano (WindowStyle = 0)
' e abre diretamente o sistema no navegador padrao.
' =============================================================================

Option Explicit

Dim fso, shell, pastaProjeto, caminhoBat

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Obtém o diretório exato onde este arquivo .vbs está salvo
pastaProjeto = fso.GetParentFolderName(WScript.ScriptFullName)
caminhoBat = fso.BuildPath(pastaProjeto, "iniciar_estoque.bat")

If fso.FileExists(caminhoBat) Then
    ' Define o diretório de trabalho atual para a pasta do projeto
    shell.CurrentDirectory = pastaProjeto
    ' O parâmetro 0 oculta totalmente a janela preta do terminal
    ' O parâmetro False libera o script VBS imediatamente após disparar o processo
    shell.Run Chr(34) & caminhoBat & Chr(34), 0, False
Else
    MsgBox "O arquivo 'iniciar_estoque.bat' não foi encontrado na pasta:" & vbCrLf & pastaProjeto, vbCritical, "Restaurante Mazon - Erro de Inicialização"
End If

Set shell = Nothing
Set fso = Nothing
