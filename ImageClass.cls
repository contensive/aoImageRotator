VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private GroupingID As Long
Private timeBetweenFrames As Double

Private lstImage As String
Private lstLink As String

Private rotSerial As Double

Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
    Exit Function
ErrorTrap:
    Call HandleError("ImageClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Public Sub Init(MainObject As Object)
    
    Set Main = MainObject
    
    Randomize
    
    rotSerial = Int(Rnd() * 1000)
    
    Exit Sub
ErrorTrap:
    Call HandleError("ImageClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    Dim CS As Long
    Dim imgWidth As Long

    GroupingID = KmaEncodeNumber(Main.GetAggrOption("Image Grouping", OptionString))
    If GroupingID = 0 Then
        GroupingID = Main.GetRecordID("Image Groupings", "Default")
        If GroupingID = 0 Then
            CS = Main.InsertCSContent("Image Groupings")
            If Main.iscsok(CS) Then
                GroupingID = Main.getcsinteger(CS, "id")
                Call Main.setcs(CS, "name", "default")
            End If
            Call Main.closecs(CS)
        End If
    End If
    timeBetweenFrames = KmaEncodeNumber(Main.GetAggrOption("Seconds to Display", OptionString))
    If timeBetweenFrames = 0 Then
        timeBetweenFrames = 3
    End If
    timeBetweenFrames = timeBetweenFrames * 1000
    imgWidth = KmaEncodeNumber(Main.GetAggrOption("Image Width", OptionString))
    
    If GroupingID <> 0 Then
        CS = Main.OpenCSContent("Image Grouping Rules", "ImageGroupingID=" & GroupingID, "", , , , "ID")
        If Main.CSOK(CS) Then
            Stream = GetScripts()
            Stream = Stream & "<a id=""link_" & rotSerial & """ href=""" & lstLink & """>"
            Stream = Stream & "<img border=""0"" id=""anim_" & rotSerial & """ src=""" & lstImage & """ "
            If imgWidth <> 0 Then
                Stream = Stream & "width=""" & imgWidth & """ "
            End If
            Stream = Stream & "></a>"
        End If
        Call Main.closecs(CS)
    End If
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("ImageClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetScripts() As String
    On Error GoTo ErrorTrap
    
    Dim retVal As String
    Dim CSRows
    Dim numCSRows As Long
    Dim Index As Long
    Dim CS As Long
    Dim SQL As String
    Dim ImageString As String
    Dim Filename As String
    Dim link As String
    
    SQL = "SELECT I.ImageFilename, I.link"
    SQL = SQL & " FROM ImageGroupingRules R, ccImages I"
    SQL = SQL & " Where (R.ImageID = I.Id)"
    SQL = SQL & " and (R.ImageGroupingID=" & GroupingID & ")"
    SQL = SQL & " and (I.ImageFilename is not null)"
    
    CS = Main.OpenCSSQL("Default", SQL)
    If Main.CSOK(CS) Then
        CSRows = Main.GetCSRows(CS)
        numCSRows = UBound(CSRows, 2)
        
        retVal = vbCrLf & "<script language=""JavaScript1.2"">" & vbCrLf
        retVal = retVal & "     var imageIndex_" & rotSerial & " = 0;" & vbCrLf
        retVal = retVal & "     var imageCount_" & rotSerial & " = " & numCSRows + 1 & ";" & vbCrLf
        retVal = retVal & "     var RotImages_" & rotSerial & " = new Array(" & numCSRows & ");" & vbCrLf
        retVal = retVal & "     var linkRotation_" & rotSerial & " = new Array(" & numCSRows & ");" & vbCrLf

        For Index = 0 To numCSRows
            link = kmaEncodeText(CSRows(1, Index))
            If link = "" Then
                link = "#"
            End If
            Filename = Main.ServerProtocol & Main.ServerHost & Main.ServerFilePath & CSRows(0, Index)
            Filename = kmaEncodeJavascript(Filename)
            retVal = retVal & "     RotImages_" & rotSerial & "[" & Index & "] = '" & Filename & "';" & vbCrLf
            retVal = retVal & "     linkRotation_" & rotSerial & "[" & Index & "] = '" & link & "';" & vbCrLf
            If Index = numCSRows Then
                lstImage = Filename
                lstLink = link
            End If
        Next

        retVal = retVal & vbCrLf
        retVal = retVal & "self.setTimeout('rotateImages_" & rotSerial & "()', " & timeBetweenFrames & ");" & vbCrLf
        retVal = retVal & vbCrLf

        retVal = retVal & "function rotateImages_" & rotSerial & "() {" & vbCrLf
        retVal = retVal & "     document.images[""anim_" & rotSerial & """].src = RotImages_" & rotSerial & "[imageIndex_" & rotSerial & "];" & vbCrLf
        retVal = retVal & "     document.getElementById(""link_" & rotSerial & """).href = linkRotation_" & rotSerial & "[imageIndex_" & rotSerial & "];" & vbCrLf
        retVal = retVal & "     imageIndex_" & rotSerial & "++;" & vbCrLf
        retVal = retVal & "     if (imageIndex_" & rotSerial & " == imageCount_" & rotSerial & ") {" & vbCrLf
        retVal = retVal & "         imageIndex_" & rotSerial & " = 0;" & vbCrLf
        retVal = retVal & "     }" & vbCrLf
        retVal = retVal & "     self.setTimeout(""rotateImages_" & rotSerial & "()"", " & timeBetweenFrames & ");" & vbCrLf
        retVal = retVal & "}" & vbCrLf
        
        retVal = retVal & "</script>" & vbCrLf
        
    End If
    Call Main.closecs(CS)
    
    GetScripts = retVal
    
    Exit Function
ErrorTrap:
    Call HandleError("ImageClass", "GetScripts", Err.Number, Err.Source, Err.Description, True, False)
    End Function



