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
    Init
    DebugInit
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
        For level = 1 To 6
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

Sub CalcPoints(thisLine As Long)
    
    Dim thisLevel As Variant
    Dim thisChildren As Variant
    
    thisLevel = dataTable.Cells(thisLine, cLevel).Value
    thisChildren = dataTable.Cells(thisLine, cChildren).Value
    
    If thisChildren = "" Then
        ' 子どもがいない。その場合、自身のInput値から作れば良い。
        Dim inputLower As Variant
        Dim inputUpper As Variant
        Dim inputActual As Variant
        Dim inputConsumed As Variant
        
        Dim calcedLower As Variant
        Dim calcedUpper As Variant
        Dim calcedActual As Variant
        Dim calcedConsumed As Variant
        

        GetPointBlock _
            RangeInputPointBlock(thisLine), _
            inputLower, inputUpper, inputActual, inputConsumed
        
        If inputActual = "" Then
            calcedLower = inputLower
            calcedUpper = inputUpper
            calcedActual = (inputLower + inputUpper) / 2
        Else
            calcedLower = inputActual
            calcedUpper = inputActual
            calcedActual = inputActual
        End If
        calcedConsumed = inputConsumed
        
        SetPointBlock _
            RangeLevelPointBlock(thisLine, CLng(thisLevel)), _
            calcedLower, calcedUpper, calcedActual, calcedConsumed

    Else
        ' こどもがいる場合
        Dim childrenArray() As String
        Dim childLine As Variant
        Dim sumPoint(1 To PointBlockLen) As Variant
        Dim sumRange As Range
        
        childrenArray = Split(thisChildren, ",")
        For Each childLine In childrenArray
        
            CalcPoints CLng(childLine) ' 再帰的に子どもの値を計算
            Dim childPointRange As Range
            Dim childPoint() As Variant
            Set childPointRange = RangeLevelPointBlock(CLng(childLine), CLng(thisLevel) + 1)
            childPoint = GetPointBlock2(childPointRange)
            
            ' 子どもの合計値を計算
            sumPoint(idxLower) = sumPoint(idxLower) + childPoint(idxLower)
            sumPoint(idxUpper) = sumPoint(idxUpper) + childPoint(idxUpper)
            sumPoint(idxActual) = sumPoint(idxActual) + childPoint(idxActual)
            sumPoint(idxConsumed) = sumPoint(idxConsumed) + childPoint(idxConsumed)
            
            
        Next childLine
        
        ' 子どもの合計値を書き込み
        Set sumRange = RangeLevelPointBlock(CLng(thisLine), CLng(thisLevel))
        SetPointBlock2 sumRange, sumPoint
        
    End If
    
End Sub


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

