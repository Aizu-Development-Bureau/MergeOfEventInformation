Sub GetDataFromWorkbooks()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim startPasteRow As Long
    Dim dataRange As Range
    
    ' �o�͐惏�[�N�V�[�g��ݒ�
    Set destWs = ThisWorkbook.Sheets(1)
    
    ' �t�H���_�̃p�X���w��
    folderPath = ThisWorkbook.Path & "\data\"
   
    ' �o�͐�t�@�C�����ƃt�@�C���p�X���w��
    fileResultName = ThisWorkbook.Path & "\result.xlsm" ' �V�����t�@�C������ݒ�

    ' �t�H���_���̍ŏ���Excel�t�@�C���̖��O���擾
    fileName = Dir(folderPath & "*.xls*")
    
    ' �t�H���_���̂��ׂĂ�Excel�t�@�C�������[�v����
    Do While fileName <> ""
        ' Excel�t�@�C�����J��
        Set wb = Workbooks.Open(folderPath & fileName)
        
        ' �t�@�C�����̃��[�N�V�[�g��ݒ�
        Set ws = wb.Sheets(1)
        
        ' �ŏI�s���擾
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        
        ' �f�[�^�����݂���ꍇ�ɂ̂ݏ��������s
        If lastRow >= 8 Then ' 8�s�ڂ���f�[�^���J�n���邽��
            ' �o�͐�̎��̋󔒍s���擾�i�擪5�s���󂯂�j
            startPasteRow = destWs.Cells(destWs.Rows.Count, 2).End(xlUp).Row + 1
            If startPasteRow < 6 Then startPasteRow = 6 ' �Œ�ł�6�s�ڂ���f�[�^���J�n����
            
            ' �f�[�^�͈͂�ݒ�iNo��������j
            Set dataRange = ws.Range("B8:M" & lastRow)
            
            ' �f�[�^���R�s�[���ē\��t����
            dataRange.Copy
            destWs.Cells(startPasteRow, 2).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        End If
        
        ' �t�@�C�������
        wb.Close SaveChanges:=False
        
        ' �t�H���_���̎���Excel�t�@�C���̖��O���擾
        fileName = Dir
    Loop
    
    ' �\�[�g���s�����߂̍ŏI�s���擾
    lastRow = destWs.Cells(destWs.Rows.Count, 2).End(xlUp).Row

    ' B��i���t��j����Ƀf�[�^�������Ń\�[�g����
    With destWs.Sort
        .SortFields.Clear
        .SortFields.Add Key:=destWs.Range("B7:B" & lastRow), Order:=xlAscending
        .SetRange destWs.Range("A7:M" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' �ʂ��ԍ���ǉ�
    For i = 8 To lastRow
        destWs.Cells(i, "A").Value = i - 7 ' 7�s�ڂ���ʂ��ԍ���U��n�߂�
    Next i
    
    Call FormatTimeColumn()

    ' �t�@�C����ʖ��ŕۑ�
    ThisWorkbook.SaveCopyAs Filename:= fileResultName

    ' �f�[�^���N���A����͈͂���肷��
    ' �����ł�B��̃f�[�^����ɂ��čŏI�s�������Ă��܂�
    lastRow = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row
    
    ' �o�̓f�[�^�͈̔͂��N���A����
    destWs.Range("B8:M" & lastRow).ClearContents


    ' Unicode�̔g�_�b�V�����g�p���āAD��̎w�肳�ꂽ�͈͂ɑ}��
    ' �S�p�g�_�b�V����Unicode��"U+301C"�ł����A
    ' Windows�ł͈�ʂ�"U+FF5E"���g�p���܂��B
    Dim waveDash As String
    waveDash = ChrW(&HFF5E)

    ' D���"~"����͂���
    destWs.Range("D8:D" & lastRow).Value = waveDash
End Sub

Sub FormatTimeColumn()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' �\��t���惏�[�N�V�[�g��ݒ�
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row ' ���ԃf�[�^�̍ŏI�s���擾

    ' ���ԃf�[�^���܂܂��C��̃Z���̏����𓝈ꂷ��
    With ws.Range("G8:G" & lastRow)
        .NumberFormat = "h:mm;@" ' ���Ԃ̏����� "��:��" �ɐݒ�
        .HorizontalAlignment = xlCenter ' �Z���̓��e�𒆉��񂹂ɂ���
    End With
End Sub
