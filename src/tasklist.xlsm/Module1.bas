Attribute VB_Name = "Module1"
Option Explicit ' �ϐ��錾��K�{��
Option Base 1   ' �z��̐擪��1�����

Dim taskWorksheet As Worksheet  ' �^�X�N�ꗗ�̃��[�N�V�[�g

Const MaxLevel As Long = 10  ' �K�w�[��

' �X�g�[���[�|�C���g�́A���l�A��l�A���l(���ےl)�A����l ��4�_�Z�b�g�B���̃I�t�Z�b�g
Const PointBlockLen As Long = 4
Const offsetLower = 0
Const offsetUpper = 1
Const offsetActual = 2
Const offsetConsumed = 3
Const idxLower = 1
Const idxUpper = 2
Const idxActual = 3
Const idxConsumed = 4

Dim dataTable As Range   ' �f�[�^�e�[�u��������Range�B���L�J�����ԍ��A�s�Ԓl�A�q�l�͂���Range��ł̔ԍ�

Dim cLineNum As Long                           ' �s�Ԃ̃J�����ԍ� (�v�Z�J����)
Dim cLevel As Long                             ' �K�w�̃J�����ԍ� (�v�Z�J����)
Dim cChildren As Long                          ' �q�̃J�����ԍ� (�v�Z�J����)
Dim cTasks(1 To MaxLevel) As Long              ' �^�X�N�����͗��̃J�����ԍ� (���̓J����)
Dim cAssignee As Long                          ' �S���ғ��͗��̃J�����ԍ� (���̓J����)
Dim cLevelPointBlocks(1 To MaxLevel) As Long   ' ���x���̃|�C���g�̃J�����ԍ� (�v�Z�J����)
Dim cInputPointBlock As Long                   ' ���̓|�C���g���̃J�����ԍ� (���̓J����)

Dim taskHeaderRow As Long  ' �^�X�N�ꗗ�V�[�g�̃w�b�_�s�ԍ��B���̒��ォ��f�[�^�e�[�u��
Dim taskTableTop As Long   ' �^�X�N�̃f�[�^�e�[�u�������̐擪�s
Dim taskTableBottom As Long ' �^�X�N�̃f�[�^�e�[�u���̖����s
Dim taskTableLeft As Long  ' �^�X�N�̃f�[�^�e�[�u�������̍��[�J����
Dim taskTableRight As Long  ' �^�X�N�̃f�[�^�e�[�u�������̉E�[�J����




Sub run()
    Init
    DebugInit
    
    CheckOrElseExit
    protectWithPresentAllows taskWorksheet, UserInterfaceOnly:=True


    CalcLineNum
    CalcLevel
    CalcChildren
    
    ' ''' test
    ' ''' CalcPoints 1
    CalcPointsTopLevel
    
End Sub

' �J�����ԍ���s�ԍ��ȂǁA�萔�l�̂悤�Ȃ��̂��v�Z���ăZ�b�g����
Sub Init()
    Set taskWorksheet = Worksheets("�^�X�N")

    ' �J�����ԍ��ϐ����Z�b�g����
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
    
    ' �s�ԍ��ϐ����Z�b�g����
    taskHeaderRow = 3
    taskTableTop = taskHeaderRow + 1
    taskTableBottom = CalcTableBottom
    
    Set dataTable = taskWorksheet.Range( _
        Cells(taskTableTop, taskTableLeft), _
        Cells(taskTableBottom, taskTableRight) _
    )
End Sub

' Init������̃Z�b�g���e�̃f�o�b�O
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


' �Ӑ}�������͂������̂������Ă��܂����Ƃ�����邽�߂ɁA
' �v�Z�l������Z���͕ҏW�s��ԂɂȂ��Ă��邱�Ƃ��m�F����
Sub CheckOrElseExit()
    Dim messages As New Collection
    
    ' �v�Z�l������J������ Locked �̏�Ԃ��`�F�b�N����
    If Not isColumnAllLocked(cLineNum) Then messages.Add "�s�ԃJ������Locked�łȂ�"
    If Not isColumnAllLocked(cLevel) Then messages.Add "�K�w�J������Locked�łȂ�"
    If Not isColumnAllLocked(cChildren) Then messages.Add "�q�J������Locked�łȂ�"

    Dim lvl As Integer
    For lvl = 1 To MaxLevel
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetLower) Then
           messages.Add "�K�w" & lvl & "�̉��l�J������Locked�łȂ�"
        End If
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetUpper) Then
            messages.Add "�K�w" & lvl & "�̉��l�J������Locked�łȂ�"
        End If
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetActual) Then
          messages.Add "�K�w" & lvl & "�̉��l�J������Locked�łȂ�"
        End If
        If Not isColumnAllLocked(cLevelPointBlocks(lvl) + offsetConsumed) Then
            messages.Add "�K�w" & lvl & "�̏���l�J������Locked�łȂ�"
        End If
    Next lvl
    
    ' �V�[�g���ی��Ԃ��ǂ������`�F�b�N����
    If Not isSheetProtected Then
        messages.Add "�V�[�g���ی��ԂɂȂ��Ă��Ȃ�"
    End If


    If messages.Count > 0 Then
        Dim concated As String
        concated = "���L�̗��R�ɂ�菈���𒆎~���܂��I" & Chr(10)
        
        Dim m As Variant
        For Each m In messages
            concated = concated & CStr(m) & Chr(10)
        Next m
        
        MsgBox concated
        End
    End If

End Sub

' dataTable �̎w�肵���J�����̃Z�����A�S�� Locked ��ԂɂȂ��Ă��邩��Ԃ�
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

' �V�[�g���ی��Ԃ��ǂ�����Ԃ�
Function isSheetProtected() As Boolean
    isSheetProtected = taskWorksheet.ProtectContents
End Function



' �������ێ������܂܁AProtect���������
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


' �^�X�N�V�[�g����A�f�[�^�e�[�u�������̍ŉ��s�̍s�ԍ����擾����B
' �^�X�N���͗��Œl�������Ă���ŉ��s���̗p�B
' �V�[�g�ɑ΂����΍s��Ԃ�
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

' �s�ԍ�(�f�[�^�e�[�u�����ł̒ʂ��s�ԍ�)���v�Z���AcLineNum�̃J�����ɏo�͂���
Sub CalcLineNum()
    Dim line As Long
    For line = 1 To dataTable.Rows.Count
        dataTable.Cells(line, cLineNum).Value = line
    Next line
End Sub

' �e�s�̊K�w���x�����v�Z���AcLevel �̃J�����ɏo�͂���
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

' �e�s�̎q�v�f�s���v�Z���AcChildren�̃J�����ɏo�͂���
Sub CalcChildren()

    Dim line As Long
    For line = 1 To dataTable.Rows.Count
        ' ���̍s�Ƀ^�X�N������A���A���̍s���ŉ��s�ł͂Ȃ�����
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
                        ' �����Ɠ���������ȏ�̃��x�����o���B�T���͂����܂ŁB
                        Exit For
                    ElseIf thatLevel = thisLevel + 1 Then
                        ' ����͎q�v�f�Ȃ̂Œǉ�
                        children.Add searchLine
                    Else
                        ' �����͑��ȍ~
                        If children.Count = 0 Then
                            MsgBox "�e���Ȃ������o���Bline=" & line & " searchLine=" & searchLine
                        End If
                        ' ��{�I�ɂ̓X�L�b�v
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


' �|�C���g�u���b�N�̒l���擾����
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


' �|�C���g�u���b�N�̃Z���ɏ�������
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
        ' �q�ǂ������Ȃ��B���̏ꍇ�A���g��Input�l������Ηǂ��B
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
        ' ���ǂ�������ꍇ
        Dim childrenArray() As String
        Dim childLine As Variant
        Dim sumPoint(1 To PointBlockLen) As Variant
        Dim sumRange As Range
        
        childrenArray = Split(thisChildren, ",")
        For Each childLine In childrenArray
        
            CalcPoints CLng(childLine) ' �ċA�I�Ɏq�ǂ��̒l���v�Z
            Dim childPointRange As Range
            Dim childPoint() As Variant
            Set childPointRange = RangeLevelPointBlock(CLng(childLine), CLng(thisLevel) + 1)
            childPoint = GetPointBlock2(childPointRange)
            
            ' �q�ǂ��̍��v�l���v�Z
            sumPoint(idxLower) = sumPoint(idxLower) + childPoint(idxLower)
            sumPoint(idxUpper) = sumPoint(idxUpper) + childPoint(idxUpper)
            sumPoint(idxActual) = sumPoint(idxActual) + childPoint(idxActual)
            sumPoint(idxConsumed) = sumPoint(idxConsumed) + childPoint(idxConsumed)
            
            
        Next childLine
        
        ' �q�ǂ��̍��v�l����������
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

