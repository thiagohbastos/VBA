'AUTOR: THIAGO BASTOS / git: thiagohbastos

'MACRO RESPONSÁVEL POR GERAR AVISO EM TELA DE HORA EM HORA
Option Explicit
Dim aviso_cronometro As Date

Sub Executa_por_tempo()

aviso_cronometro = Now + TimeValue("01:00:00")

Call Application.OnTime(aviso_cronometro, "executa_por_tempo")
MsgBox ("Pausa Para Água!")

End Sub

Sub finaliza_por_tempo()

Call Application.OnTime(aviso_cronometro, "executa_por_tempo", , False)

End Sub