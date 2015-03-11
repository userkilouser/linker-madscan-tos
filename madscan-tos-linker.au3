#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <Misc.au3>
#include <Array.au3>


; Ссылка на xml, содержащий название комании
Global Const $yqlAPICompanyNameRequest = "http://d.yimg.com/autoc.finance.yahoo.com/autoc?query=<SYMBOL>&callback=YAHOO.Finance.SymbolSuggest.ssCallback"

; Ссылка на текст, содержащий сектор и индустрии компании
Global Const $yqlAPICompanySectorRequest = "http://www.stocksinplay.ru/quote.php?tickersInput=<SYMBOL>"

; регистрация нажатия ESC для выхода из программы
HotKeySet("{ESC}", "Terminate")

; Создаем окно формы
$pic = GUICreate("Linker", 400, 30, 620, 80, $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_TOPMOST)) ;

; Кладем на форму картинку с прозрачным фоном
$basti_stay = GUICtrlCreatePic("bground.gif", 0, 0, 400, 30,-1, $GUI_WS_EX_PARENTDRAG)

; Создаем надпись (пока пустую)
$hDC = GUICtrlCreateLabel("",0, 0, 400, 30)
; Настройка надписи
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetColor($hDC, 0xffd800)

; Отображаем окно формы
GUISetState(@SW_SHOW)

; Инициализация тикера
$symbPrev = ""

; "Вечный" цикл отображения окна формы
While 1

   ; Берем видимый текст с активного окна
   Local $hActiveText = WinGetText("[ACTIVE]", "")

   ;ConsoleWrite("HH" & $hActiveText & @CRLF)

   ; Сравниваем полученную выше строку с известным значением WinGetText() для фильтров Madscan
   If StringInStr($hActiveText, "toolStripContainer1") = 2 Then

	  ;ConsoleWrite("MS: " & $hActiveText & @CRLF)

	  ; Обнулям предыдущее значение надписи
      ;ControlSetText($pic, "", $hDC, "")

	  ; Если активное окно - это фильтр Madscan, то посылаем ему Ctrl+C для копирования в буфер всей строки, которая под мышкой
      Send("{CTRLDOWN}C{CTRLUP}")

	  ; Убираем из строки часть из времени алерта (которое в американском формате, например 1:13 PM)
      Local $Clip = StringRegExpReplace (ClipGet(), ":\d+\s[A|P]M", "", 0)

      ; Выбираем из отстатка строки тикер
      Local $TickerArray = StringRegExp($Clip, '([A-Z|\.\-\+]+)\s', 1, 1)
      Local $Ticker = _ArrayToString($TickerArray, "")
	  ;ConsoleWrite("$TickerArray: " & $TickerArray & @CRLF)
	  ;ConsoleWrite("$Ticker: " & $Ticker & @CRLF)

	  ; Обновляем $symbPrev
	  $symbPrev = $Ticker

	  ; Активируем окно Level2 в Arche
       _WinWaitActivate("[CLASS:SunAwtFrame]", "")
      Local $hLevelII = ControlGetHandle("[CLASS:SunAwtFrame]", "", "")
	  ;ConsoleWrite("$hLevelII: " & $hLevelII & @CRLF)
	  ; ControlClick("", "", "[CLASS:SunAwtFrame]", "left", 2, 106, 66)
      ControlSend ("", "", $hLevelII, $Ticker & "{ENTER}", 0)
	  ;ConsoleWrite("@error: " & @error & @CRLF)
;~
;~ 	  For $element In $TickerArray
;~ 		 Send($element)
;~ 	  Next
;~ 	  Send( "{ENTER}")

	  ; Вызов функции для получения инфо компании по тикеру
      $sSymbolInfo = GetCompanyInfo($Ticker)

	  ; Устанавливаем значения надписи в соответствии с инфо о компании
      GUICtrlSetData($hDC, $sSymbolInfo)

   EndIf

   ; Ручное информирование :)

   ; Определяем заголовок активного окна
   ; Local $windowTitle = WinGetTitle("[ACTIVE]", "")
   ; ConsoleWrite("$windowTitle=" & $windowTitle & @LF)

   ; Если активное окно - это окно Level2, то
   ; If StringInStr($windowTitle, "Level2") = 1 Then

      ; Определяем видимый текст окна Level2
      ; Local $hActiveText = $windowTitle
      ; ConsoleWrite("$hActiveText=" & $hActiveText & @LF)

      ; Из видимого текста выбираем тикер, по паттерну
      ; $symbArray = StringRegExp($hActiveText, '([A-Z|\.\-\+]+)~', 1, 1)
      ; If @error > 0 Then
          ; ConsoleWrite("StringRegExp@error=" & @error & @LF)
          ; MsgBox(0, "StringRegExp@error", @error)
      ; EndIf

      ; Тикер
      ; Local $symbNew = _ArrayToString($symbArray, "")

      ; Если значение тикера изменилось, то
      ; If $symbNew <> $symbPrev Then
         ; ConsoleWrite("$symbNew=" & $symbNew & @LF )

         ; Обнулям предыдущее значение надписи
         ; ControlSetText($pic, "", $hDC, "")

         ; Вызов функции для получения инфо компании по тикеру
         ; $sSymbolInfo = GetCompanyInfo($symbNew)

         ; Устанавливаем значения надписи в соответствии с инфо о компании
         ; GUICtrlSetData($hDC, $sSymbolInfo)

         ; Заменяем предыдущее значение тикера на новое
         ; $symbPrev = $symbNew
      ; EndIf

   ; EndIf

   ; Если нажата правая клавиша мышки - выход из цикла
   If _IsPressed("02") Then
      ExitLoop
   EndIf

	; Снятие нагрузки с процессора
	Sleep(500)

WEnd

; Функция активации окна
Func _WinWaitActivate($title,$text,$timeout=0)
    WinWait($title,$text,$timeout)
    If Not WinActive($title,$text) Then WinActivate($title,$text)
    WinWaitActive($title,$text,$timeout)
 EndFunc

; Выход из программы
Func Terminate()
    Exit 0
 EndFunc

; Получение инфо о компании по тикеру
Func GetCompanyInfo($sSymbol)

   $sRequest = StringReplace($yqlAPICompanyNameRequest, "<SYMBOL>", $sSymbol)
   ;ConsoleWrite("$sRequest: " & $sRequest & @CRLF)

   ; Получение информации об имени компании
   $bData = InetRead($sRequest)

   $aLines = BinaryToString($bData, 4)
   $aLines = StringReplace($aLines, "},{", @CRLF)
   ;ConsoleWrite("$aLines: " & $aLines & @CRLF)

   $array = StringRegExp($aLines, '"name":(.*),"exch":.*:"(.*)",', 1, 1)

   If @error = 0 then
	  ;ConsoleWrite ("$array[0]: " & $array[0] & @CRLF)
	  ;ConsoleWrite ("$array[1]: " & $array[1] & @CRLF)
	  $sCompanyInfo = ($array[0] & ", " & $array[1] & ", " & @CRLF)
   Else
	  $sCompanyInfo = "N/A, "
   EndIf

   ; Получение информации о секторе и индустрии компании

   $sRequest = StringReplace($yqlAPICompanySectorRequest, "<SYMBOL>", $sSymbol)
   ; ConsoleWrite($sRequest & @CRLF)
   $bData = InetRead($sRequest)

   $aLines = BinaryToString($bData, 4)
   ConsoleWrite("$aLines-bs" & $aLines & @CRLF)

   $array = StringRegExp($aLines, 'panel-title">(.*, )(.*,)( .*,)( .*)</h3', 1, 1)
   If @error = 0 then
      ;ConsoleWrite ("0: " & $array[0] & @CRLF)
      ;ConsoleWrite ("1: " & $array[1] & @CRLF)
	  ;ConsoleWrite ("2: " & $array[2] & @CRLF)
	  ;ConsoleWrite ("3: " & $array[3] & @CRLF)
      $sCompanyInfo = $sCompanyInfo & $array[1] & $array[2] & $array[3]
   Else
	  $sCompanyInfo = $sCompanyInfo & "N/A"
   EndIf

   Return $sCompanyInfo

EndFunc