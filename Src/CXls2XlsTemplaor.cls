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
'簡易的なテンプレートクラスです。１つのクラスで完結できるように設計されています。
'
'BaseCellが左上の基準セルになります。たとえばA1を指定した場合、A1を基準にテンプレートを反映します。
'出力先シートのクリアは行いません。適時クリアしてください。処理に失敗した場合、ロールバックを行い元の状態に戻ります。
'テンプレートのヘッダー部分・アイテム部分・フッター部分はすべて１つのシートに設定してください。
'なお、ヘッダーおよびフッターは設定しなくても構いません。その場合、左上基準からアイテム領域の先頭が反映されます。
'アイテム領域でデータを設定する列はデータソースのレコードセットのフィールド名を設定してください。
'データソースとなるレコードセットはカーソルロケーションをadUseClientに設定してください。
'Copyright (c) 2016 TskGunGun
'Released under the MIT license
'https://github.com/TskGunGun/
'
'本ソフトウェアは自由な変更を認めます。また、コピー及び再配布・商用利用など自由に使ってください。
'再配布する場合、本ライセンス全文をソースコード、もしくはライセンス表示ファイルなどに掲載してください。

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
