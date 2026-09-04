Option Explicit

' ============================================================
' ГЛОБАЛЬНЫЙ ФЛАГ ОТМЕНЫ
' ============================================================

Public CancelRequested As Boolean

' Проверяем ABORT реже для большей скорости
Private AbortCounter As Long

' Увеличено для лучшей производительности
Private Const ABORT_CHECK_EVERY As Long = 100


' ============================================================
' ПРОВЕРКА КНОПКИ ABORT
' ============================================================

Private Function CheckAbort(Optional ByVal ForceCheck As Boolean = False) As Boolean

    AbortCounter = AbortCounter + 1

    If ForceCheck Or AbortCounter >= ABORT_CHECK_EVERY Then

        AbortCounter = 0

        DoEvents

        If CancelRequested Then
            CheckAbort = True
            Exit Function
        End If

    End If

    CheckAbort = False

End Function


' ============================================================
' ОПТИМИЗАЦИЯ: КЭШИРОВАНИЕ ГРАНИЦ
' ============================================================

Private Function IsShapeInBounds(s As Shape, _
                                  leftBound As Double, _
                                  rightBound As Double, _
                                  topBound As Double, _
                                  bottomBound As Double) As Boolean

    IsShapeInBounds = (s.CenterX >= leftBound And _
                       s.CenterX <= rightBound And _
                       s.CenterY >= bottomBound And _
                       s.CenterY <= topBound)

End Function


' ============================================================
' ОПТИМИЗАЦИЯ: БЫСТРЫЙ РАСЧЁТ ПЛОЩАДИ
' ============================================================

Private Function GetShapeArea(s As Shape) As Double

    GetShapeArea = s.SizeWidth * s.SizeHeight

End Function


' ============================================================
' ОСНОВНОЙ МАКРОС - ОПТИМИЗИРОВАННЫЙ
' ============================================================

Sub PrepareStandsWithInputBox()

    Dim s As Shape
    Dim innerShape As Shape
    Dim checkStand As Shape
    Dim compareShape As Shape

    Dim lr As Layer

    Dim initialSelection As ShapeRange

    Dim stands As New Collection
    Dim candidates As New Collection

    Dim i As Long
    Dim j As Long

    Dim shapeCount As Long
    Dim candidateCount As Long

    Dim isStand As Boolean
    Dim isInsideAnother As Boolean

    Dim leftBoundary As Double
    Dim rightBoundary As Double
    Dim topBoundary As Double
    Dim bottomBoundary As Double

    Dim maxAreaInSelection As Double
    Dim areaThreshold As Double
    Dim tempArea As Double

    Dim subIdx As Long
    Dim sp As SubPath

    Dim standsCount As Long

    Dim userInput As String

    Dim isRedSelected As Boolean
    Dim isGreenSelected As Boolean

    Dim msgPrompt As String
    Dim msgTitle As String
    Dim msgSuccess As String

    Dim maxSubArea As Double
    Dim currentSubArea As Double
    Dim mainSubPathIndex As Long

    Dim duplicateFound As Boolean
    Dim uniqueStands As New Collection

    Dim candidatesArray() As Shape
    Dim candidateAreas() As Double
    
    Dim innerCenterX As Double
    Dim innerCenterY As Double
    Dim standLeftX As Double
    Dim standRightX As Double
    Dim standBottomY As Double
    Dim standTopY As Double
    Dim subPathCount As Long


    ' --------------------------------------------------------
    ' СБРОС ABORT
    ' --------------------------------------------------------

    CancelRequested = False
    AbortCounter = 0


    ' --------------------------------------------------------
    ' ПРОВЕРКА ДОКУМЕНТА
    ' --------------------------------------------------------

    If ActiveDocument Is Nothing Then Exit Sub


    ' --------------------------------------------------------
    ' ПРОВЕРКА ВЫДЕЛЕНИЯ
    ' --------------------------------------------------------

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select objects first!", vbExclamation
        Exit Sub
    End If


    On Error GoTo ErrorHandler


    ' --------------------------------------------------------
    ' НАЧАЛО КОМАНДЫ
    ' --------------------------------------------------------

    ActiveDocument.BeginCommandGroup "Clean Stands"


    ' --------------------------------------------------------
    ' ОКНО ABORT
    ' --------------------------------------------------------

    frmAbort.lblStatus.Caption = "Preparing..."
    frmAbort.cmdAbort.Enabled = True

    frmAbort.Show vbModeless

    DoEvents


    ' --------------------------------------------------------
    ' НАСТРОЙКА
    ' --------------------------------------------------------

    ActiveDocument.Unit = cdrMillimeter

    Set lr = ActivePage.ActiveLayer


    ' --------------------------------------------------------
    ' ЗАПОМИНАЕМ ГРАНИЦЫ ИСХОДНОГО ВЫДЕЛЕНИЯ
    ' --------------------------------------------------------

    Set initialSelection = ActiveSelectionRange

    leftBoundary = initialSelection.LeftX
    rightBoundary = initialSelection.RightX
    topBoundary = initialSelection.TopY
    bottomBoundary = initialSelection.BottomY


    ' ========================================================
    ' РАЗБЛОКИРОВКА
    ' ========================================================

    frmAbort.lblStatus.Caption = "Unlocking objects..."

    shapeCount = lr.Shapes.Count

    For i = 1 To shapeCount

        If i Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = lr.Shapes(i)
        If s.Locked Then s.Locked = False

    Next i


    ' ========================================================
    ' ПОЛНОЕ РАЗГРУППИРОВАНИЕ
    ' ========================================================

    frmAbort.lblStatus.Caption = "Ungrouping..."

    If CheckAbort(True) Then GoTo AbortMacro

    initialSelection.UngroupAll

    If CheckAbort(True) Then GoTo AbortMacro

    ActiveDocument.ClearSelection


    ' ========================================================
    ' ОБНОВЛЯЕМ КОЛИЧЕСТВО ОБЪЕКТОВ
    ' ========================================================

    shapeCount = lr.Shapes.Count


    ' ========================================================
    ' ИЩЕМ МАКСИМАЛЬНУЮ ПЛОЩАДЬ
    ' ========================================================

    frmAbort.lblStatus.Caption = "Analyzing objects..."

    maxAreaInSelection = 0

    For i = 1 To shapeCount

        If i Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = lr.Shapes(i)

        If IsShapeInBounds(s, leftBoundary, rightBoundary, topBoundary, bottomBoundary) Then

            tempArea = GetShapeArea(s)

            If tempArea > maxAreaInSelection Then
                maxAreaInSelection = tempArea
            End If

        End If

    Next i


    If maxAreaInSelection <= 0 Then
        MsgBox "No suitable objects found!", vbExclamation
        GoTo NormalExit
    End If


    areaThreshold = maxAreaInSelection / 15


    ' ========================================================
    ' РАЗБИВАЕМ МЕЛКИЕ КРИВЫЕ
    ' ========================================================

    frmAbort.lblStatus.Caption = "Breaking small curves..."

    shapeCount = lr.Shapes.Count

    For i = shapeCount To 1 Step -1

        If i Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = lr.Shapes(i)

        If IsShapeInBounds(s, leftBoundary, rightBoundary, topBoundary, bottomBoundary) Then

            If s.Type = cdrCurveShape And GetShapeArea(s) < areaThreshold Then

                On Error Resume Next
                s.BreakApart
                On Error GoTo ErrorHandler

            End If

        End If

    Next i


    shapeCount = lr.Shapes.Count


    ' ========================================================
    ' СОБИРАЕМ ТОЛЬКО КРУПНЫЕ КАНДИДАТЫ
    ' ========================================================

    frmAbort.lblStatus.Caption = "Finding stand contours..."

    Set candidates = New Collection

    For i = 1 To shapeCount

        If i Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = lr.Shapes(i)

        If IsShapeInBounds(s, leftBoundary, rightBoundary, topBoundary, bottomBoundary) Then

            If GetShapeArea(s) >= areaThreshold Then
                candidates.Add s
            End If

        End If

    Next i


    candidateCount = candidates.Count


    ' ========================================================
    ' КОНВЕРТИРУЕМ В МАССИВ ДЛЯ СКОРОСТИ
    ' ========================================================

    If candidateCount > 0 Then

        ReDim candidatesArray(1 To candidateCount)
        ReDim candidateAreas(1 To candidateCount)

        For i = 1 To candidateCount
            Set candidatesArray(i) = candidates(i)
            candidateAreas(i) = GetShapeArea(candidatesArray(i))
        Next i

    End If


    ' ========================================================
    ' ИЩЕМ НАРУЖНЫЕ РАМКИ
    ' ========================================================

    For i = 1 To candidateCount

        If i Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = candidatesArray(i)
        isInsideAnother = False

        For j = 1 To candidateCount

            If candidateAreas(j) > candidateAreas(i) Then

                Set checkStand = candidatesArray(j)

                If s.CenterX >= checkStand.LeftX And _
                   s.CenterX <= checkStand.RightX And _
                   s.CenterY >= checkStand.BottomY And _
                   s.CenterY <= checkStand.TopY Then

                    isInsideAnother = True
                    Exit For

                End If

            End If

        Next j

        If Not isInsideAnother Then
            stands.Add s
        End If

    Next i


    If stands.Count = 0 Then
        MsgBox "No contours found!", vbExclamation
        GoTo NormalExit
    End If


    ' ========================================================
    ' УДАЛЯЕМ ДУБЛИКАТЫ
    ' ========================================================

    frmAbort.lblStatus.Caption = "Removing duplicate contours..."

    Set uniqueStands = New Collection

    For i = 1 To stands.Count

        If i Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = stands(i)
        duplicateFound = False

        For j = 1 To uniqueStands.Count

            Set compareShape = uniqueStands(j)

            If Abs(s.CenterX - compareShape.CenterX) < 0.05 And _
               Abs(s.CenterY - compareShape.CenterY) < 0.05 And _
               Abs(s.SizeWidth - compareShape.SizeWidth) < 0.05 And _
               Abs(s.SizeHeight - compareShape.SizeHeight) < 0.05 Then

                duplicateFound = True
                Exit For

            End If

        Next j

        If Not duplicateFound Then
            uniqueStands.Add s
        Else
            On Error Resume Next
            s.Delete
            On Error GoTo ErrorHandler
        End If

    Next i


    Set stands = uniqueStands
    standsCount = stands.Count

    If standsCount = 0 Then
        MsgBox "No contours found!", vbExclamation
        GoTo NormalExit
    End If


    ' ========================================================
    ' УДАЛЯЕМ ВСЁ ВНУТРИ РАМОК
    ' ========================================================

    frmAbort.lblStatus.Caption = "Removing inner objects..."

    shapeCount = lr.Shapes.Count

    For i = shapeCount To 1 Step -1

        If i Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set innerShape = lr.Shapes(i)

        isStand = False

        ' Проверяем — не является ли объект самой рамкой
        For j = 1 To stands.Count
            If innerShape.StaticID = stands(j).StaticID Then
                isStand = True
                Exit For
            End If
        Next j

        If Not isStand Then

            If IsShapeInBounds(innerShape, leftBoundary, rightBoundary, topBoundary, bottomBoundary) Then

                innerCenterX = innerShape.CenterX
                innerCenterY = innerShape.CenterY

                For j = 1 To stands.Count

                    Set s = stands(j)

                    standLeftX = s.LeftX
                    standRightX = s.RightX
                    standBottomY = s.BottomY
                    standTopY = s.TopY

                    If innerCenterX >= standLeftX And _
                       innerCenterX <= standRightX And _
                       innerCenterY >= standBottomY And _
                       innerCenterY <= standTopY Then

                        innerShape.Delete
                        Exit For

                    End If

                Next j

            End If

        End If

    Next i


    ' ========================================================
    ' ОБРАБОТКА SUBPATHS
    ' ========================================================

    frmAbort.lblStatus.Caption = "Cleaning contours..."

    For j = 1 To stands.Count

        If j Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = stands(j)

        If s.Type <> cdrCurveShape Then
            s.ConvertToCurves
        End If

        maxSubArea = 0
        mainSubPathIndex = 1
        subPathCount = s.Curve.SubPaths.Count

        For subIdx = 1 To subPathCount

            Set sp = s.Curve.SubPaths(subIdx)
            currentSubArea = sp.BoundingBox.Width * sp.BoundingBox.Height

            If currentSubArea > maxSubArea Then
                maxSubArea = currentSubArea
                mainSubPathIndex = subIdx
            End If

        Next subIdx

        For subIdx = subPathCount To 1 Step -1
            If subIdx <> mainSubPathIndex Then
                s.Curve.SubPaths(subIdx).Delete
            End If
        Next subIdx

    Next j


    standsCount = stands.Count


    ' ========================================================
    ' ЗАКРЫВАЕМ ОКНО ОБРАБОТКИ
    ' ========================================================

    Unload frmAbort

    DoEvents


    ' ========================================================
    ' ВЫБОР ЦВЕТА
    ' ========================================================

    msgPrompt = "Total prepared: " & standsCount & vbCrLf & vbCrLf & _
                "R - RED" & vbCrLf & _
                "G - GREEN"

    msgTitle = "Color"

    userInput = InputBox(msgPrompt, msgTitle, "R")
    userInput = UCase(userInput)

    If userInput = "" Then
        GoTo NormalExit
    End If


    ' ========================================================
    ' ПРОВЕРКА ЦВЕТА
    ' ========================================================

    isRedSelected = (userInput = "R") Or _
                    (userInput = Chr(202)) Or _
                    (userInput = Chr(234))

    isGreenSelected = (userInput = "G") Or _
                      (userInput = Chr(207)) Or _
                      (userInput = Chr(239))


    ' ========================================================
    ' ФИНАЛЬНЫЙ ОКРАС
    ' ========================================================

    CancelRequested = False
    AbortCounter = 0

    frmAbort.lblStatus.Caption = "Applying color..."
    frmAbort.cmdAbort.Enabled = True

    frmAbort.Show vbModeless

    DoEvents

    For j = 1 To stands.Count

        If j Mod ABORT_CHECK_EVERY = 0 And CheckAbort Then GoTo AbortMacro

        Set s = stands(j)

        With s
            .Fill.ApplyNoFill
            .Outline.Width = 0.076

            If isRedSelected Then
                .Outline.Color.CMYKAssign 0, 100, 100, 0
            ElseIf isGreenSelected Then
                .Outline.Color.CMYKAssign 100, 0, 100, 0
            End If

        End With

    Next j


    ' ========================================================
    ' НОРМАЛЬНОЕ ЗАВЕРШЕНИЕ
    ' ========================================================

NormalExit:

    On Error Resume Next
    Unload frmAbort
    ActiveDocument.EndCommandGroup
    On Error GoTo 0

    If standsCount > 0 Then
        msgSuccess = "Success! Contours combined: " & standsCount
        MsgBox msgSuccess, vbInformation, "CorelDRAW"
    End If

    Exit Sub


    ' ========================================================
    ' ABORT
    ' ========================================================

AbortMacro:

    On Error Resume Next
    Unload frmAbort
    ActiveDocument.ClearSelection
    ActiveDocument.EndCommandGroup
    On Error GoTo 0

    MsgBox "Operation aborted by user.", vbExclamation, "Clean Stands"

    Exit Sub


    ' ========================================================
    ' ERROR
    ' ========================================================

ErrorHandler:

    On Error Resume Next
    Unload frmAbort
    ActiveDocument.EndCommandGroup
    On Error GoTo 0

    MsgBox "Error: " & Err.Description, vbCritical, "CorelDRAW"

End Sub
