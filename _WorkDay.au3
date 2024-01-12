; ----------------------------------------------------------------------------
; Autor: Rodrigo lamoglia Vitorino
; E-mail: rlvitorin@gmail.com
; Data de Cria��o: 12 de janeiro de 2024
; Descri��o: Este script em AutoIt realiza v�rias opera��es, incluindo a obten��o
; de datas �teis, verifica��o de feriados e muito mais.
; ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>
#include <JSON.au3>
#include <Array.au3>
#include <Date.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>

; Cria vetor de feriados
Global $aHolidays[0]

; URL da API
Local $sUrl = "https://brasilapi.com.br/api/feriados/v1/" & @YEAR

; Diret�rio e nome do arquivo INI para salvar as datas processadas
Local $sIniFile = "c:\temp\ProcessedDates.ini"
;~ Local $sIniFile = @ScriptDir & "\ProcessedDates.ini"

; Cria o objeto XMLHTTP
Local $oHTTP = ObjCreate("Microsoft.XMLHTTP")

; Verifica se o objeto foi criado com sucesso
If IsObj($oHTTP) Then

	; Define a solicita��o GET
    $oHTTP.open("GET", $sUrl, False)
    $oHTTP.send()

    ; Verifica se a solicita��o foi bem-sucedida (status 200)
    If $oHTTP.status = 200 Then

		; Resposta da API em formato JSON
        Local $sResponse = $oHTTP.responseText

        ; Voc� pode processar a resposta aqui conforme necess�rio
        ; MsgBox($MB_OK, "Resposta da API", $sResponse)

        $o_Object = _JSON_Parse($sResponse)

        For $i = 0 To UBound($o_Object) - 1 Step 1
            ; MsgBox("","",(($o_Object[$i])["date"]))
            _ArrayAdd($aHolidays, (($o_Object[$i])["date"]))
        Next
    Else
        MsgBox($MB_OK, "Erro na solicita��o", "Ocorreu um erro ao fazer a solicita��o. C�digo de status: " & $oHTTP.status)
    EndIf

    ; Libera o objeto XMLHTTP
    $oHTTP = 0
Else
    MsgBox($MB_OK, "Erro", "Falha ao criar o objeto Microsoft.XMLHTTP")
EndIf

; Fun��o para verificar se uma data � fim de semana (s�bado ou domingo)
Func IsWeekend($sDate)
    Local $iDayOfWeek = _DateToDayOfWeek(GetYear($sDate), GetMonth($sDate), GetDay($sDate))
    Return $iDayOfWeek == 7 Or $iDayOfWeek == 1 ; S�bado ou domingo
EndFunc

; Fun��o para verificar se uma data � feriado
Func IsHoliday($sDate)
	$sDate = StringReplace($sDate,"/","-")
    Return _ArraySearch($aHolidays, $sDate) > -1
EndFunc

; Fun��o para obter o ano de uma data
Func GetYear($sDate)
    Local $aDate = StringSplit($sDate, "/")
    If UBound($aDate) = 4 Then
        Return $aDate[1]
    EndIf
    Return ""
EndFunc

; Fun��o para obter o m�s de uma data
Func GetMonth($sDate)
    Local $aDate = StringSplit($sDate, "/")
    If UBound($aDate) = 4 Then
        Return $aDate[2]
    EndIf
    Return ""
EndFunc

; Fun��o para obter o dia de uma data
Func GetDay($sDate)
    Local $aDate = StringSplit($sDate, "/")
    If UBound($aDate) = 4 Then
        Return $aDate[3]
    EndIf
    Return ""
EndFunc

; Fun��o para calcular o pr�ximo dia �til
Func GetNextWorkday($sToday)
    While True
        $sNextDay = _DateAdd("d", 1, $sToday)
        Local $iNextDayOfWeek = _DateToDayOfWeek(GetYear($sNextDay), GetMonth($sNextDay), GetDay($sNextDay))
        If (Not IsWeekend($sNextDay)) And (Not IsHoliday($sNextDay)) Then
            Return $sNextDay
        EndIf
        $sToday = $sNextDay
    WEnd
EndFunc

; Fun��o para verificar se uma data j� foi processada
Func IsDateProcessed($sDate, $sIniFile)
    Return IniRead($sIniFile, "ProcessedDates", $sDate, "")
EndFunc

; Fun��o para marcar uma data como processada
Func MarkDateProcessed($sDate, $sIniFile)
    IniWrite($sIniFile, "ProcessedDates", $sDate, "Processed")
EndFunc

; Fun��o para obter o primeiro dia �til do m�s
Func GetFirstWorkdayOfMonth($sYear, $sMonth)
    Local $sFirstDayOfMonth = $sYear & "/" & $sMonth & "/01"
    While True
        If Not IsWeekend($sFirstDayOfMonth) And Not IsHoliday($sFirstDayOfMonth) Then
            Return $sFirstDayOfMonth
        EndIf
        $sFirstDayOfMonth = _DateAdd("d", 1, $sFirstDayOfMonth)
    WEnd
EndFunc

; Fun��o para obter o �ltimo dia �til do m�s
Func GetLastWorkdayOfMonth($sYear, $sMonth)
    Local $sFirstDayOfNextMonth = _DateAdd("M", 1, $sYear & "/" & $sMonth & "/01")
    Local $sLastDayOfMonth = _DateAdd("d", -1, $sFirstDayOfNextMonth)
    While True
        If Not IsWeekend($sLastDayOfMonth) And Not IsHoliday($sLastDayOfMonth) Then
            Return $sLastDayOfMonth
        EndIf
        $sLastDayOfMonth = _DateAdd("d", -1, $sLastDayOfMonth)
    WEnd
EndFunc

;Fun��o para obter ultimo dia do m�s
Func GetLastDayOfMonth($sYear, $sMonth)
    Local $sLastDayOfMonth
    If $sMonth = 2 And Mod($sYear, 4) = 0 And (Mod($sYear, 100) <> 0 Or Mod($sYear, 400) = 0) Then
        ; Fevereiro em ano bissexto
        $sLastDayOfMonth = "29/" & $sMonth & "/" & $sYear
    Else
        $sLastDayOfMonth = _DateAdd("d", -1, _DateAdd("M", 1, $sYear & "/" & $sMonth & "/01"))
    EndIf
    Return $sLastDayOfMonth
EndFunc

; Obter as datas dos feriados do m�s
Local $sToday = _NowCalcDate()
Local $sFirstDayOfMonth = @YEAR & "/" & @MON & "/01"

If $sToday > $sFirstDayOfMonth Then
	$sToday = $sFirstDayOfMonth
EndIf

; Loop para processar os dias �teis do m�s
While True
    If Not IsWeekend($sToday) And Not IsHoliday($sToday) And Not IsDateProcessed($sToday, $sIniFile) Then

		; Execute o c�digo que deseja aqui
		ConsoleWrite("Processando dia:" & $sToday & @CRLF)

        ; Marque a data como processada no INI
        MarkDateProcessed($sToday, $sIniFile)

    EndIf

    ; Avance para o pr�ximo dia �til
    $sToday = GetNextWorkday($sToday)

	; Se Processamento D-1 operador compara��o � igual a "=", sen�o at� dia atual igual a ">"
	If $sToday >  _NowCalcDate() Then
		ExitLoop
		;Exit
	EndIf

    ; Se o pr�ximo dia �til for no pr�ximo m�s, saia do loop
    If GetMonth($sToday) <> @MON Then

		ExitLoop
		;Exit

    EndIf
WEnd

; Exemplo de uso:
Local $sYear = @YEAR
Local $sMonth = @MON
Local $sToday = _NowCalcDate()

Local $sFirstWorkday = GetFirstWorkdayOfMonth($sYear, $sMonth)
Local $sLastWorkday = GetLastWorkdayOfMonth($sYear, $sMonth)
Local $sLastDayOfMonth = GetLastDayOfMonth($sYear, $sMonth)
Local $sIsDateProcessed = IsDateProcessed($sToday, $sIniFile)
Local $sNextWorkday = GetNextWorkday($sToday)

ConsoleWrite("Primeiro dia �til do m�s: " & $sFirstWorkday & @CRLF)
ConsoleWrite("�ltimo dia �til do m�s: " & $sLastWorkday & @CRLF)
ConsoleWrite("�ltimo dia do m�s: " & $sLastDayOfMonth & @CRLF)
ConsoleWrite("Data j� processada ? : " & $sIsDateProcessed  & " >> Date:" & $sToday & @CRLF)
ConsoleWrite("Pegar pr�ximo dia �til : " & $sNextWorkday  & " >> Data Refer�ncia:" & $sToday & @CRLF)
ConsoleWrite("Hoje � Feriado ? : " & IsHoliday($sToday) & @CRLF)
ConsoleWrite("� final de semana ? : " & IsWeekend($sToday) & @CRLF)