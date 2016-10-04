Attribute VB_Name = "Module1"
Option Explicit ' 変数宣言を必須に
Option Base 1   ' 配列の先頭を1からに

Dim taskWorksheet As Worksheet  ' タスク一覧のワークシート

Const MaxLevel As Long = 10  ' 階層深さ

' ストーリーポイントは、下値、上値、仮値(実際値)、消費値 の4点セット。そのオフセット
Const PointBlockLen As Long = 4
Const offsetLower = 0
Const offsetUpper = 1
Const offsetActual = 2
Const offsetConsumed = 3
Const idxLower = 1
Const idxUpper = 2
Const idxActual = 3
Const idxConsumed = 4

Dim dataTable As Range   ' データテーブル部分のRange。下記カラム番号、行番値、子値はこのRange基準での番号

Dim cLineNum As Long                           ' 行番のカラム番号 (計算カラム)
Dim cLevel As Long                             ' 階層のカラム番号 (計算カラム)
Dim cChildren As Long                          ' 子のカラム番号 (計算カラム)
Dim cTasks(1 To MaxLevel) As Long              ' タスク名入力欄のカラム番号 (入力カラム)
Dim cAssignee As Long                          ' 担当者入力欄のカラム番号 (入力カラム)
Dim cLevelPointBlocks(1 To MaxLevel) As Long   ' レベルのポイントのカラム番号 (計算カラム)
Dim cInputPointBlock As Long                   ' 入力ポイント欄のカラム番号 (入力カラム)

Dim taskHeaderRow As Long  ' タスク一覧シートのヘッダ行番号。この直後からデータテーブル
Dim taskTableTop As Long   ' タスクのデータテーブル部分の先頭行
Dim taskTableBottom As Long ' タスクのデータテーブルの末尾行
Dim taskTableLeft As Long  ' タスクのデータテーブル部分の左端カラム
Dim taskTableRight As Long  ' タスクのデータテーブル部分の右端カラム




Sub run()
    Application.ScreenUpdating = False   ' 高速化のため
    
    Init
    DebugInit
    
    CheckOrElseExit
    protectWithPresentAllows taskWorksheet, UserInterfaceOnly:=True
    
    clearCalculationArea
    CalcLineNum
    CalcLevel
    CalcChildren
    
    ' ''' test
    ' ''' CalcPoints 1
    CalcPointsTopLevel
    
End Sub

' カラム番号や行番号など、定数値のようなものを計算してセットする
Sub Init()
    Set taskWorksheet = Worksheets("タスク")

    ' カラム番号変数をセットする
    Dim c As Long
    Dim level As Long
    c = 1
    cLineNum = c: c = c + 1
    cLevel = c: c = c + 1
    cChildren = c: c = c + 1
    For level = 1 To MaxLevel
        cTasks(level) = c: c = c + 1
    Next level
    cAssignee = c: c = c + 1
    
    For level = 1 To MaxLevel
        cLevelPointBlocks(level) = c + PointBlockLen * (level - 1)
    Next level
    cInputPointBlock = c + PointBlockLen * MaxLevel
    
    taskTableLeft = cLineNum
    taskTableRight = cInputPointBlock + PointBlockLen - 1
    
    ' 行番号変数をセットする
    taskHeaderRow = 3
    taskTableTop = taskHeaderRow + 1
    taskTableBottom = CalcTableBottom
    
    Set dataTable = taskWorksheet.Range( _
        Cells(taskTableTop, taskTableLeft), _
        Cells(taskTableBottom, taskTableRight) _
    )
End Sub

' Init処理後のセット内容のデバッグ
Sub DebugInit()
    Dim level As Long

    Debug.Print "===="
    Debug.Print "taskWorksheet = " & taskWorksheet.Name
    Debug.Print "MaxLevel = " & MaxLevel
    Debug.Print "cLineNum = " & cLineNum
    Debug.Print "cLevel = " & cLevel

    Debug.Print "cChildren = " & cChildren
    For level = 1 To MaxLevel
        Debug.Print "cTasks(" & level & ") = " & cTasks(level)
    Next level
    Debug.Print "cAssignee = " & cAssignee
    For level = 1 To MaxLevel
        Debug.Print "cLevelPointBlocks(" & level & ") = " & cLevelPointBlocks(level)
    Next level
    Debug.Print "cInputPointBlock = " & cInputPointBlock
    
    Debug.Print "taskHeaderRow = " & taskHeaderRow
    Debug.Print "taskTableTop = " & taskTableTop
    Debug.Print "taskTableBottom = " & taskTableBottom
    Debug.Print "taskTableLeft = " & taskTableLeft
    Debug.Print "taskTableRight = " & taskTableRight
    
    Debug.Print "dataTable = " & dataTable.Address
End Sub


' 意図せず入力したものを消してしまうことを避けるために、
' 計算値が入るセルは編集不可状態になっていることを確認する
Sub CheckOrElseExit()
    Dim messages As New Collection
    
    ' 計算値が入るカラムの Locked の状態をチェックする
    If Not isColumnAllLocked(cLineNum) Then messages.Add "行番カラムがLockedでない"
    If Not isColumnAllLocked(cLevel) Then messages.Add "階層カラムがLockedでない"
    If Not isColumnAllLocked(cChildren) Then messages.Add "子カラムがLockedでない"

    Dim lvl As Integer
    For lvl = 1 To MaxLevel
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetLower) Then
           messages.Add "階層" & lvl & "の下値カラムがLockedでない"
        End If
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetUpper) Then
            messages.Add "階層" & lvl & "の下値カラムがLockedでない"
        End If
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetActual) Then
          messages.Add "階層" & lvl & "の仮値カラムがLockedでない"
        End If
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetConsumed) Then
            messages.Add "階層" & lvl & "の消費値カラムがLockedでない"
        End If
    Next lvl
    
    ' シートが保護状態かどうかをチェックする
    If Not isSheetProtected Then
        messages.Add "シートが保護状態になっていない"
    End If


    If messages.Count > 0 Then
        Dim concated As String
        concated = "下記の理由により処理を中止します！" & Chr(10)
        
        Dim m As Variant
        For Each m In messages
            concated = concated & CStr(m) & Chr(10)
        Next m
        
        MsgBox concated
        End
    End If

End Sub

' dataTable の指定したカラムのセルが、全て Locked 状態になっているかを返す
Function isColumnAllLocked(c As Long) As Boolean

    Dim status As Variant
    status = dataTable.Columns(c).Locked
    ' Debug.Print "status = " & TypeName(status)
    
    If IsNull(status) Then
        isColumnAllLocked = False
    Else
        isColumnAllLocked = CBool(status)
    End If

End Function

' シートが保護状態かどうかを返す
Function isSheetProtected() As Boolean
    isSheetProtected = taskWorksheet.ProtectContents
End Function



' 許可操作を保持したまま、Protect操作をする
Sub protectWithPresentAllows( _
    sheet As Worksheet, _
    Optional UserInterfaceOnly As Boolean = False)

    Dim p As Object
    Set p = sheet.Protection

    sheet.Protect _
        UserInterfaceOnly:=UserInterfaceOnly, _
        AllowFormattingCells:=p.AllowFormattingCells, _
        AllowFormattingColumns:=p.AllowFormattingColumns, _
        AllowFormattingRows:=p.AllowFormattingRows, _
        AllowInsertingColumns:=p.AllowInsertingColumns, _
        AllowInsertingRows:=p.AllowInsertingRows, _
        AllowInsertingHyperlinks:=p.AllowInsertingHyperlinks, _
        AllowDeletingColumns:=p.AllowDeletingColumns, _
        AllowDeletingRows:=p.AllowDeletingRows, _
        AllowSorting:=p.AllowSorting, _
        AllowFiltering:=p.AllowFiltering, _
        AllowUsingPivotTables:=p.AllowUsingPivotTables

End Sub


' 計算領域をクリアする
Sub clearCalculationArea()
    dataTable.Columns(cLineNum).ClearContents
    dataTable.Columns(cLevel).ClearContents
    dataTable.Columns(cChildren).ClearContents
    dataTable.Columns(cChildren).ClearContents

    Dim c As Integer
    For c = cLevelPointBlocks(1) To cLevelPointBlocks(MaxLevel) + offsetConsumed
        dataTable.Columns(c).ClearContents
    Next c

End Sub

Function RangeLevelPointBlock(line As Long, level As Long) As Range
    Set RangeLevelPointBlock = _
        dataTable.Range( _
            Cells(line, cLevelPointBlocks(level)), _
            Cells(line, cLevelPointBlocks(level) + offsetConsumed) _
        )
End Function

Function RangeInputPointBlock(line As Long) As Range
    Set RangeInputPointBlock = _
        dataTable.Range( _
            Cells(line, cInputPointBlock), _
            Cells(line, cInputPointBlock + offsetConsumed) _
        )
End Function


' タスクシートから、データテーブル部分の最下行の行番号を取得する。
' タスク入力欄で値が入っている最下行を採用。
' シートに対する絶対行を返す
Function CalcTableBottom() As Long
    Dim column As Variant
    Dim bottom As Long
    Dim maxRows As Long
    
    maxRows = taskWorksheet.Rows.Count
    bottom = 0
    For Each column In cTasks
        Dim tmp As Long
        tmp = Cells(maxRows, column).End(xlUp).Row
        If bottom < tmp Then bottom = tmp
    Next column
    CalcTableBottom = bottom
End Function

' 行番号(データテーブル内での通し行番号)を計算し、cLineNumのカラムに出力する
Sub CalcLineNum()
    Dim line As Long
    For line = 1 To dataTable.Rows.Count
        dataTable.Cells(line, cLineNum).Value = line
    Next line
End Sub

' 各行の階層レベルを計算し、cLevel のカラムに出力する
Sub CalcLevel()
    
    Dim line As Long
    For line = 1 To dataTable.Rows.Count
        
        Dim level As Long
        For level = 1 To MaxLevel
            If dataTable.Cells(line, cTasks(level)).Value <> "" Then Exit For
        Next level
        
        If level > MaxLevel Then
            dataTable.Cells(line, cLevel).Value = Empty
        Else
            dataTable.Cells(line, cLevel).Value = level
        End If
    
    Next line
End Sub

' 各行の子要素行を計算し、cChildrenのカラムに出力する
Sub CalcChildren()

    Dim line As Long
    For line = 1 To dataTable.Rows.Count
        ' その行にタスクがあり、かつ、その行が最下行ではないこと
        If dataTable.Cells(line, cLevel).Value <> "" And line < dataTable.Rows.Count Then
        
            Dim children As Collection
            Dim searchLine As Long
            Dim thisLevel As Long

            thisLevel = dataTable.Cells(line, cLevel).Value
            Set children = New Collection
            
            For searchLine = line + 1 To dataTable.Rows.Count
                Dim thatLevel As Variant
                thatLevel = dataTable.Cells(searchLine, cLevel).Value
                If thatLevel <> "" Then
                    If thatLevel <= thisLevel Then
                        ' 自分と同じかそれ以上のレベルが出現。探索はそこまで。
                        Exit For
                    ElseIf thatLevel = thisLevel + 1 Then
                        ' これは子要素なので追加
                        children.Add searchLine
                    Else
                        ' ここは孫以降
                        If children.Count = 0 Then
                            MsgBox "親がなく孫が出現。line=" & line & " searchLine=" & searchLine
                        End If
                        ' 基本的にはスキップ
                    End If
                End If
            Next searchLine
            
            If children.Count > 0 Then
                Dim s As String
                Dim i As Long
                s = ""
                For i = 1 To children.Count
                    If i > 1 Then s = s & ","
                    s = s & children(i)
                Next i
                dataTable(line, cChildren).Value = s
            Else
                dataTable(line, cChildren).Value = Empty
            End If
        
        End If
    Next line
End Sub


' ポイントブロックの値を取得する
Sub GetPointBlock( _
    block As Range, _
    lower As Variant, _
    upper As Variant, _
    actual As Variant, _
    consumed As Variant)
    
    lower = block.Cells(1, 1 + offsetLower).Value
    upper = block.Cells(1, 1 + offsetUpper).Value
    actual = block.Cells(1, 1 + offsetActual).Value
    consumed = block.Cells(1, 1 + offsetConsumed).Value
End Sub
    
Function GetPointBlock2(r As Range) As Variant
    Dim point(1 To PointBlockLen) As Variant
    
    point(idxUpper) = r.Item(idxUpper).Value
    point(idxLower) = r.Item(idxLower).Value
    point(idxActual) = r.Item(idxActual).Value
    point(idxConsumed) = r.Item(idxConsumed).Value
    
    GetPointBlock2 = point
End Function


' ポイントブロックのセルに書き込む
Sub SetPointBlock( _
    block As Range, _
    lower As Variant, _
    upper As Variant, _
    actual As Variant, _
    consumed As Variant)

    block.Cells(1, 1 + offsetLower).Value = lower
    block.Cells(1, 1 + offsetUpper).Value = upper
    block.Cells(1, 1 + offsetActual).Value = actual
    block.Cells(1, 1 + offsetConsumed).Value = consumed
End Sub

Sub SetPointBlock2(r As Range, point() As Variant)
    r.Item(idxLower).Value = point(idxLower)
    r.Item(idxUpper).Value = point(idxUpper)
    r.Item(idxActual).Value = point(idxActual)
    r.Item(idxConsumed).Value = point(idxConsumed)
End Sub

' 指定行から、子要素について再帰的に計算し、最終的に指定行の値をセルにセット、
' 関数の返り値として返す。
Function CalcPoints(thisLine As Long) As Variant
    
    Dim thisLevel As Variant
    Dim thisChildren As Variant
    
    thisLevel = dataTable.Cells(thisLine, cLevel).Value
    thisChildren = dataTable.Cells(thisLine, cChildren).Value
    
    Dim inputPoint() As Variant
    Dim calcedPoint() As Variant
    
    inputPoint = GetPointBlock2(RangeInputPointBlock(thisLine))
    
    
    If thisChildren = "" Then
        ' 子どもがいない。その場合、自身のInput値から作れば良い。
        
        calcedPoint = CalcPointLogicNoChild(inputPoint)
        
    Else
        ' 子どもがいる場合
        Dim childrenArray() As String
        Dim childLine As Variant
        
        Dim childrenPoints As New Collection
        
        childrenArray = Split(thisChildren, ",")
        For Each childLine In childrenArray
            Dim childPoint() As Variant
            childPoint = CalcPoints(CLng(childLine)) ' 再帰的に子どもの値を計算
            childrenPoints.Add Item:=childPoint      ' 結果を childrenPoints に貯めていく
        Next childLine
                
        calcedPoint = CalcPointLogicWithChildren(inputPoint, childrenPoints)
        
    End If
    
    SetPointBlock2 _
        RangeLevelPointBlock(thisLine, CLng(thisLevel)), _
        calcedPoint
        
    CalcPoints = calcedPoint
    
End Function

Function CalcPointLogicNoChild( _
    inputPoint() As Variant _
) As Variant

    Dim calcedPoint(1 To PointBlockLen) As Variant

    If inputPoint(idxActual) = "" Then
        If Not IsEmpty(inputPoint(idxLower)) Or Not IsEmpty(inputPoint(idxUpper)) Then
           calcedPoint(idxLower) = inputPoint(idxLower)
            calcedPoint(idxUpper) = inputPoint(idxUpper)
           calcedPoint(idxActual) = (inputPoint(idxLower) + inputPoint(idxUpper)) / 2
        End If
    Else
        calcedPoint(idxLower) = inputPoint(idxActual)
        calcedPoint(idxUpper) = inputPoint(idxActual)
        calcedPoint(idxActual) = inputPoint(idxActual)
    End If
    calcedPoint(idxConsumed) = inputPoint(idxConsumed)
    
    CalcPointLogicNoChild = calcedPoint

End Function


Function CalcPointLogicWithChildren( _
    inputPoint() As Variant, _
    childrenPoints As Collection _
) As Variant

    Dim calcedPoint() As Variant

    Dim sum_of_delta_lower_square As Variant
    Dim sum_of_delta_upper_square As Variant
    Dim sum_of_middle As Variant
    Dim sum_of_consumed As Variant

    Dim childPoint As Variant
    For Each childPoint In childrenPoints

        ' 子どもの合計値を計算 (２乗和平方根法:SRSS法:Square Root Sum of Squares)
        If Not IsEmpty(childPoint(idxLower)) Or Not IsEmpty(childPoint(idxUpper)) _
                Or Not IsEmpty(childPoint(idxActual)) Or Not IsEmpty(childPoint(idxConsumed)) Then
            sum_of_delta_lower_square = sum_of_delta_lower_square + (childPoint(idxActual) - childPoint(idxLower)) ^ 2
            sum_of_delta_upper_square = sum_of_delta_upper_square + (childPoint(idxUpper) - childPoint(idxActual)) ^ 2
            sum_of_middle = sum_of_middle + childPoint(idxActual)
            sum_of_consumed = sum_of_consumed + childPoint(idxConsumed)
        End If

    Next childPoint

    If Not IsEmpty(sum_of_delta_lower_square) Or Not IsEmpty(sum_of_delta_upper_square) _
            Or Not IsEmpty(sum_of_middle) Or Not IsEmpty(sum_of_consumed) Then

        ReDim calcedPoint(1 To PointBlockLen)
        calcedPoint(idxLower) = sum_of_middle - Sqr(sum_of_delta_lower_square)
        calcedPoint(idxUpper) = sum_of_middle + Sqr(sum_of_delta_upper_square)
        calcedPoint(idxActual) = sum_of_middle
        calcedPoint(idxConsumed) = sum_of_consumed

    Else
        calcedPoint = CalcPointLogicNoChild(inputPoint)
    End If

    CalcPointLogicWithChildren = calcedPoint
End Function



Sub CalcPointsTopLevel()
    Dim line As Long
    For line = 1 To dataTable.Rows.Count
        If dataTable.Cells(line, cLevel) = 1 Then
            CalcPoints line
        End If
    Next line
End Sub



Sub test()
    Dim block As Variant
    block = RangeInputPointBlock(3).Value
    
    Debug.Print block(idxLower)
    Debug.Print block(idxUpper)
    Debug.Print block(idxActual)
    Debug.Print block(idxConsumed)

End Sub

