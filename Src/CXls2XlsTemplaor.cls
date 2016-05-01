VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CXls2XlsTemplaor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CXls2XlsTemplator
'�ȈՓI�ȃe���v���[�g�N���X�ł��B�P�̃N���X�Ŋ����ł���悤�ɐ݌v����Ă��܂��B
'
'BaseCell������̊�Z���ɂȂ�܂��B���Ƃ���A1���w�肵���ꍇ�AA1����Ƀe���v���[�g�𔽉f���܂��B
'�o�͐�V�[�g�̃N���A�͍s���܂���B�K���N���A���Ă��������B�����Ɏ��s�����ꍇ�A���[���o�b�N���s�����̏�Ԃɖ߂�܂��B
'�e���v���[�g�̃w�b�_�[�����E�A�C�e�������E�t�b�^�[�����͂��ׂĂP�̃V�[�g�ɐݒ肵�Ă��������B
'�Ȃ��A�w�b�_�[����уt�b�^�[�͐ݒ肵�Ȃ��Ă��\���܂���B���̏ꍇ�A��������A�C�e���̈�̐擪�����f����܂��B
'�A�C�e���̈�Ńf�[�^��ݒ肷���̓f�[�^�\�[�X�̃��R�[�h�Z�b�g�̃t�B�[���h����ݒ肵�Ă��������B
'�f�[�^�\�[�X�ƂȂ郌�R�[�h�Z�b�g�̓J�[�\�����P�[�V������adUseClient�ɐݒ肵�Ă��������B
'Copyright (c) 2016 TskGunGun
'Released under the MIT license
'https://github.com/TskGunGun/
'
'�{�\�t�g�E�F�A�͎��R�ȕύX��F�߂܂��B�܂��A�R�s�[�y�эĔz�z�E���p���p�Ȃǎ��R�Ɏg���Ă��������B
'�Ĕz�z����ꍇ�A�{���C�Z���X�S�����\�[�X�R�[�h�A�������̓��C�Z���X�\���t�@�C���ȂǂɌf�ڂ��Ă��������B

'------------------------------------------------
'Class Variables
'------------------------------------------------
'Template Sheet Setting
Private TemplateSheet As Worksheet
Private DirSheet As Worksheet
Private SrcData As ADODB.Recordset
 
'BasePoint Setting
Private BaseCell_Row As Integer
Private BaseCell_Column As Integer

'Template Range Setting
Public HeaderRange As Range
Public ItemRange As Range
Public FooterRange As Range
'------------------------------------------------
'Class Procedures
'------------------------------------------------
'Constructor
Private Sub Class_initialize()
    Set TemplateSheet = Nothing
    Set DirSheet = Nothing
    Set HeaderRange = Nothing
    Set ItemRange = Nothing
    Set FooterRange = Nothing
    Set SrcData = Nothing
    BaseCell_Row = 0
    BaseCell_Column = 0
End Sub

'Destructor
Private Sub Class_terminate()
    Set SrcData = Nothing
End Sub

'------------------------------------------------
'Public Procedures
'------------------------------------------------
Public Sub SetRange(Header As Range, Item As Range, Footer As Range)
    Set HeaderRange = Header
    Set ItemRange = Item
    Set FooterRange = Footer
End Sub

Public Sub SetBaseCellIndex(RowIndex As Long, ColIndex As Long)
    BaseCell_Row = RowIndex
    BaseCell_Column = ColIndex
End Sub

Public Sub SetDataSource(ByRef Src As ADODB.Recordset)
    Set SrcData = Src
End Sub

Public Sub SetTemplateSheet(ByRef Template As Worksheet)
    Set TemplateSheet = Template
End Sub

Public Sub SetDirSheet(ByRef Dir As Worksheet)
    Set DirSheet = Dir
End Sub

'Return Value is Success : TRUE Failure:FALSE
Public Function Execute() As Boolean
    Execute = False
    If Not ValidationParam Then Exit Function
    
    On Error GoTo ensure:
    Dim work As Worksheet
    Set work = ActiveWorkbook.Worksheets.Add
    
    If Not HeaderRange Is Nothing Then CopyHeader work
    
    CopyItem work
    
    If Not FooterRange Is Nothing Then CopyFooter work
    
    MsgBox "Finish", vbOKOnly + vbInformation

ensure:
    'work.Delete
End Function

'------------------------------------------------
'Private Procedures
'------------------------------------------------[
Private Sub CopyHeader(ByRef work As Worksheet)
    Dim BaseCell As Range
    Set BaseCell = Range(Cells(BaseCell_Row, BaseCell_Column).Address)
    CopyRange TemplateSheet, HeaderRange, work, BaseCell
End Sub

Private Sub CopyItem(ByRef work As Worksheet)

    Dim RowIndex As Integer
    RowIndex = BaseCell_Row + HeaderRange.Rows.Count
    
    Dim ll As Integer
    Dim Fields() As String
    For ll = 0 To SrcData.Fields.Count - 1
        ReDim Preserve Fields(ll)
        Fields(ll) = SrcData.Fields(ll).Name
    Next
    
    Dim ii As Integer
    For ii = 0 To SrcData.RecordCount - 1
        CopyRange TemplateSheet, ItemRange, work, Range(Cells(RowIndex + ii, BaseCell_Column).Address)
        
        Dim jj As Integer
        For jj = 0 To UBound(Fields)
            work.Cells(RowIndex + ii, BaseCell_Column + jj).Value = SrcData.Fields(Fields(jj)).Value
        Next
        
        SrcData.MoveNext
    Next
    
End Sub

Private Sub CopyFooter(ByRef work As Worksheet)
    Dim FooterBase_Row As Long
    FooterBase_Row = BaseCell_Row + HeaderRange.Rows.Count + SrcData.RecordCount
    Dim BaseCell As Range
    Set BaseCell = Range(Cells(FooterBase_Row, BaseCell_Column).Address)
    CopyRange TemplateSheet, FooterRange, work, BaseCell
End Sub

Private Sub CopyRange(ByRef SrcSheet As Worksheet, SrcRange As Range, _
                                       ByRef DirSheet As Worksheet, BasePoint As Range)

    Application.ScreenUpdating = False
    Dim ActiveSheetName As String: ActiveSheetName = ActiveSheet.Name
    SrcSheet.Activate
    SrcSheet.Range(SrcRange.Address).Select
    Selection.Copy
    
    DirSheet.Activate
    DirSheet.Range(BasePoint.Address).Select
    DirSheet.Paste
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

'Validation of Param, Return Value is Success: TRUE Failure: FALSE
Private Function ValidationParam() As Boolean
    ValidationParam = Not TemplateSheet Is Nothing
    ValidationParam = ValidationParam And Not DirSheet Is Nothing
    ValidationParam = ValidationParam And Not ItemRange Is Nothing
    
    ValidationParam = ValidationParam And Not SrcData Is Nothing
    ValidationParam = ValidationParam And SrcData.RecordCount > 0
    
    ValidationParam = ValidationParam And BaseCell_Row <> 0
    ValidationParam = ValidationParam And BaseCell_Column <> 0
End Function

Public Sub ArrayRemove(ByRef TargetArray As Variant, ByVal deleteIndex As Integer)
    Dim i As Integer

    For i = deleteIndex To UBound(TargetArray) - 1
        TargetArray(i) = TargetArray(i + 1)
    Next i

    ReDim Preserve TargetArray(UBound(TargetArray) - 1)

End Sub