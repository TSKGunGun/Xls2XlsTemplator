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
'https://github.com/TSKGunGun/Xls2XlsTemplator
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

'Other Setting
Public ColumnWide As Boolean

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
    ColumnWide = True
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
    
    On Error GoTo catch:
    Dim work As Worksheet
    Set work = ActiveWorkbook.Worksheets.Add
    
    On Error GoTo ensure:
    If Not HeaderRange Is Nothing Then CopyHeader work
    
    CopyItem work
    If ColumnWide Then ModColumnWides
    
    If Not FooterRange Is Nothing Then CopyFooter work
    
    CopyRange work, GetTemplateRange, DirSheet, GetBaseCell
    Execute = True
    Debug.Print "CXls2XlsTemplator's Execute Succes!"
ensure:
    DeleteSheet work
    Exit Function

catch:
    Debug.Print "CXls2XlsTemplator's Execute Failled!"
End Function

'------------------------------------------------
'Private Procedures
'------------------------------------------------[
Private Sub CopyHeader(ByRef work As Worksheet)
    CopyRange TemplateSheet, HeaderRange, work, GetBaseCell
End Sub

Private Sub CopyItem(ByRef work As Worksheet)

    Dim RowIndex As Integer
    RowIndex = BaseCell_Row + HeaderRange.Rows.Count
    
    CopyRange TemplateSheet, ItemRange, work, Cells(RowIndex, BaseCell_Column), SrcData.RecordCount
    Dim ii As Integer
    For ii = 0 To SrcData.RecordCount - 1
        Dim jj As Integer
        For jj = 0 To ItemRange.Columns.Count
            Dim InputFldName As String
            InputFldName = work.Cells(RowIndex + ii, BaseCell_Column + jj).Value
            
            If InputFldName <> "" Then
                work.Cells(RowIndex + ii, BaseCell_Column + jj).Value = SrcData.Fields(InputFldName)
            End If
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
                                       ByRef DirSheet As Worksheet, BasePoint As Range, Optional Count As Integer = 1)

    Dim ActiveSheetName As String: ActiveSheetName = ActiveSheet.Name
    SrcSheet.Activate
    SrcSheet.Range(SrcRange.Address).Select
    Selection.Copy
    
    DirSheet.Activate
    
    Dim ii As Integer
    For ii = 0 To Count - 1
        DirSheet.Range(Cells(BasePoint.Row + ii, BasePoint.Column).Address).Select
        DirSheet.Paste
    Next
    Application.CutCopyMode = False
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

Private Function GetTemplateRange() As Range
    
    Dim RowCount, ColumnCount As Integer
    RowCount = SrcData.RecordCount
    ColumnCount = SrcData.Fields.Count
    
    If Not HeaderRange Is Nothing Then
        RowCount = RowCount + HeaderRange.Rows.Count
        If HeaderRange.Columns.Count > ColumnCount Then _
            ColumnCount = HeaderRange.Columns.Count
    End If
            
    If Not FooterRange Is Nothing Then
        RowCount = RowCount + FooterRange.Rows.Count
          If FooterRange.Columns.Count > ColumnCount Then _
            ColumnCount = FooterRange.Columns.Count
    End If
    
    Set GetTemplateRange = GetBaseCell.Resize(RowCount, ColumnCount)
 
End Function

Private Function GetBaseCell() As Range
    Set GetBaseCell = Range(Cells(BaseCell_Row, BaseCell_Column).Address)
End Function

Private Sub DeleteSheet(ByRef DeleteSheet As Worksheet)
    Application.DisplayAlerts = False
    DeleteSheet.Delete
    Debug.Print "CXls2XlsTemplator's Worksheet Delete Completed"
    Application.DisplayAlerts = True
End Sub

Private Sub ModColumnWides()
    Dim ColumnIndex As Integer
    
    Dim DirCol, SrcCol As Range
    Set DirCol = DirSheet.Range(GetTemplateRange.Columns.Address)
    Set SrcCol = TemplateSheet.Range(ItemRange.Columns.Address)
    For ColumnIndex = 1 To DirCol.Columns.Count
        DirCol(ColumnIndex).ColumnWidth = SrcCol(ColumnIndex).ColumnWidth
    Next
End Sub

