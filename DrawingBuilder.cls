VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DrawingBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==== SET THESE PATHS TO VALID TEMPLATE PATHS ====='
Const drawTemplatePath As String = "E:\Mikes Data\SOLIDWORKS Stuff\Templates\Drawings\C-Size Assy.DRWDOT"
Const TableTemplate As String = "E:\Mikes Data\SOLIDWORKS Stuff\Templates\Tables\BOM Table.sldbomtbt"

Public swDraw As DrawingDoc
Private viewConfigName As String
Private firstView As View
Private swBOMAnnotation As BomTableAnnotation

Public Sub CreateDrawing(swApp As SldWorks.SldWorks)
    Set swDraw = swApp.NewDocument(drawTemplatePath, 0, 0, 0)
End Sub

Public Function SetViews(modelPath As String) As Boolean
    If swDraw Is Nothing Then
        SetViews = False
        Exit Function
    End If
    SetViews = swDraw.InsertModelInPredefinedView(modelPath)
End Function

Public Function GetFirstView(drawModel As DrawingDoc) As View
    Dim v As View
    Set v = drawModel.GetFirstView 'get the drawing sheet
    Set v = v.GetNextView
    Set GetFirstView = v
    Set firstView = v
    viewConfigName = v.ReferencedConfiguration
End Function

Private Function GetDrawModel() As ModelDoc2
    Set GetDrawModel = swDraw
End Function

Private Sub SelectFirstView()  
    Dim vName As String
    vName = firstView.Name    
    Call GetDrawModel.extension.SelectByID2(vName, "DRAWINGVIEW", 0, 0, 0, False, -1, Nothing, 0)    
End Sub


Public Function InsertBOM() As Boolean    
    Dim swView As View
    Set swView = GetFirstView(swDraw)
    Dim AnchorType As Long
    AnchorType = swBOMConfigurationAnchor_BottomRight
    Dim BomType As Long
    BomType = swBomType_TopLevelOnly
    
    Set swBOMAnnotation = swView.InsertBomTable2(True, 0.4, 0.3, AnchorType, BomType, viewConfigName, TableTemplate)
    Dim swBOMFeature As BomFeature
    Set swBOMFeature = swBOMAnnotation.BomFeature
End Function

Public Function InsertBalloons() As Boolean    
    Dim bOpt As AutoBalloonOptions
    Set bOpt = swDraw.CreateAutoBalloonOptions
    
    bOpt.Style = swBalloonLayoutType_e.swDetailingBalloonLayout_Square
    bOpt.InsertMagneticLine = True
    SelectFirstView
    swDraw.AutoBalloon5 bOpt    
End Function

Public Function SaveDrawing(newPath As String) As Boolean
    SaveDrawing = GetDrawModel.SaveAs(newPath)
End Function

Public Function SaveBOM(bomPath As String) As Boolean
    SaveBOM = swBOMAnnotation.SaveAsExcel(bomPath, False, True)
End Function


