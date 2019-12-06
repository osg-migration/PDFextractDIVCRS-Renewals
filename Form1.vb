Imports System.Xml
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports pdftron
Imports pdftron.PDF
Imports pdftron.Common
Imports System.Collections.Specialized

Public Class Form1
    Public FRMW As FRMW.FRMW
    Dim DOCU As DOCU.DOCU
    Dim convLog As ConversionLog.ConversionLog

    Dim swDOCU As StreamWriter
    Dim swSLCT As StreamWriter

    Dim CurrentPage As Page
    Dim clipAccountNumber, clipNA, clipDueDate, clipQESP As New Rect
    Dim HelveticaRegularFont As PDF.Font

    Dim nameAddressList, remittanceAddressList As New StringCollection
    Dim XObjects As Dictionary(Of String, Element)

    Dim clientCode, CurrentPDFFileName, documentID, accountNumber, workDir, QESP, pieceID, prevPieceID As String
    Dim docNumber, currentPageNumber, origPageNumber, docPageCount, StartingPage, totalPages, pageTotal As Integer
    Dim foreignAddress As Boolean
    Dim cancelledFlag As Boolean = False

    Structure TextAndStyle
        Public text As String
        Public fontName As String
        Public fontSIze As Double
    End Structure

#Region "Form Events"

    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim oRAngle As System.Drawing.Rectangle = New System.Drawing.Rectangle(0, 0, Me.Width, Me.Height)
        Dim oGradientBrush As Brush = New Drawing.Drawing2D.LinearGradientBrush(oRAngle, Color.WhiteSmoke, Color.Crimson, Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal)
        e.Graphics.FillRectangle(oGradientBrush, oRAngle)
    End Sub

    Private Sub Form1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If (Environment.ExitCode = 0) And FRMW.parse("{NormalTermination}") <> "YES" Then
            cancelledFlag = True
            Throw New Exception("Program was cancelled while executing")
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Interval = 1000
        Timer1.Enabled = True
        status("Starting")
    End Sub

    Private Sub status(ByVal txt As String)
        lblStatus.Text = txt
        Me.Refresh()
        Application.DoEvents()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        standardProcessing()
    End Sub

#End Region

#Region "OSG Process"

    Private Sub standardProcessing()
        Dim licenseKey As String
        Dim docuFileName As String

        FRMW = New FRMW.FRMW
        lblEXE.Text = Application.ExecutablePath
        FRMW.StandardInitialization(Application.ExecutablePath)
        convLog = New ConversionLog.ConversionLog("PDFextractDIVCRS - Renewals")
        DOCU = New DOCU.DOCU

        FRMW.loadFrameworkApplicationConfigFile("PDFEXTRACT")
        licenseKey = FRMW.getJOBparameter("PDFTRONLICENSELEY")
        docuFileName = FRMW.getParameter("PDFEXTRACT.outputDOCUfile")
        CurrentPDFFileName = FRMW.getParameter("PDFEXTRACT.inputPDFfile")
        clientCode = FRMW.getParameter("CLIENTCODE")
        workDir = FRMW.getParameter("WORKDIR")
        swSLCT = New StreamWriter(FRMW.getParameter("PDFEXTRACT.OUTPUTSLCTFILE"), False, Encoding.Default)


        swDOCU = New StreamWriter(docuFileName, False, Encoding.Default)
        PDFNet.Initialize(licenseKey)

        SetParsingCoordinates()

        ProcessPDF()

        swDOCU.Flush() : swDOCU.Close()
        swSLCT.Flush() : swSLCT.Close()
        PDFNet.Terminate()
        convLog.ZIPandCopy()
        FRMW.StandardTermination()

        Application.Exit()

    End Sub

    Private Sub ProcessPDF()
        ClearValues()

        Using inDoc As New PDFDoc(CurrentPDFFileName)
            pageTotal = inDoc.GetPageCount

            LoadFonts(inDoc)

            While currentPageNumber < pageTotal
                currentPageNumber += 1 'Current page number will increment as blank pages and backer are added
                origPageNumber += 1

                CurrentPage = inDoc.GetPage(currentPageNumber)
                ProcessPage(inDoc)

            End While

            'Write DOCU record for last account
            writeDOCUrecord(totalPages)

            status("Processing PDF page (" & origPageNumber.ToString & "); Saving Output PDF...")
            inDoc.Save(FRMW.getParameter("PDFExtract.OutputPDFFile"), SDF.SDFDoc.SaveOptions.e_compatibility + SDF.SDFDoc.SaveOptions.e_remove_unused)

        End Using

    End Sub

    Private Sub ClearValues()
        accountNumber = ""
        nameAddressList = New StringCollection
        totalPages = 0 : docPageCount = 1
    End Sub

    Private Sub ProcessPage(ByRef inDoc As PDFDoc)
        Dim seq As String = ""

        QESP = GetPDFpageValue(clipQESP)
        If QESP.Contains(":") Then
            pieceID = QESP.Split(":"c)(2)
            seq = QESP.Split(":"c)(3)
        End If

        'Remove 2-D bar code
        WhiteOutContentBox(0, 8, 0.45, 1, , , , 1)

        If origPageNumber Mod 100 = 0 Then
            status("Processing PDF page (" & origPageNumber.ToString & ")")
        End If

        If pieceID <> prevPieceID Then
            If Integer.Parse(seq) = 1 Then
                'Start of document will have sequence number = 1
                ProcessPage1(inDoc)
                prevPieceID = pieceID
            End If
        End If

        prevPieceID = pieceID
        totalPages += 1
        docPageCount += 1

    End Sub

    Private Sub ProcessPage1(ByRef inDoc As PDFDoc)
        If docNumber > 0 Then
            'Write DOCU record
            writeDOCUrecord(totalPages)
            ClearValues()
        End If

        'Get important values
        accountNumber = GetPDFpageValue(clipAccountNumber).Trim()
        nameAddressList = GetPDFpageValues(clipNA)
        StartingPage = currentPageNumber

        Select Case QESP.Split(":"c)(0).Replace("(QESP)", "")
            Case "04", "05", "06"
                foreignAddress = True
            Case Else
                foreignAddress = False
        End Select

        documentID = Guid.NewGuid.ToString
        CreateSLCTentry()

        'Check values
        If accountNumber = "" Then Throw New Exception(convLog.addError("Account number not found", accountNumber, "123456789", "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & currentPageNumber))
        If nameAddressList.Count = 0 Then Throw New Exception(convLog.addError("No name and address found", , , "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))

        'White out address box
        WhiteOutContentBox(0.6, 0.85, 3.75, 1, , , , 1)

        AdjustPagePosition(CurrentPage, 0, -0.25)

        docNumber += 1

    End Sub

    Private Sub CreateSLCTentry()
        Dim SLCT As New SLCT.SLCT
        SLCT.documentId = documentID
        SLCT.applicationCode = FRMW.getParameter("SCSapplicationCode")
        SLCT.accountNumber = accountNumber
        SLCT.target = ""
        SLCT.addValue("savingsCode", QESP.Substring(3, 4))
        swSLCT.WriteLine(SLCT.SLCTrecord())
        SLCT = Nothing

    End Sub


#End Region

#Region "Standard PDF Procedures"

    Private Sub SetParsingCoordinates()

        clipAccountNumber.x1 = I2P(3.02)
        clipAccountNumber.y1 = I2P(0.05)
        clipAccountNumber.x2 = (clipAccountNumber.x1 + I2P(1))
        clipAccountNumber.y2 = (clipAccountNumber.y1 + I2P(0.175))

        clipNA.x1 = I2P(0.6)
        clipNA.y1 = I2P(0.85)
        clipNA.x2 = (clipNA.x1 + I2P(3.75))
        clipNA.y2 = (clipNA.y1 + I2P(1))

        clipQESP.x1 = I2P(4.9)
        clipQESP.y1 = I2P(0.05)
        clipQESP.x2 = (clipQESP.x1 + I2P(3))
        clipQESP.y2 = (clipQESP.y1 + I2P(0.15))

        CreateCropPage()

    End Sub

    Private Sub CreateCropPage()

        Using cropDoc As New PDFDoc
            Dim page As Page = cropDoc.PageCreate(New Rect(0, 0, 612, 792))
            cropDoc.PageInsert(cropDoc.GetPageIterator(0), page)
            page = cropDoc.GetPage(1)

            'Remove x1 value from x2 for crop box creation
            CreateCropBox("ACCOUNT NUMBER", clipAccountNumber.x1, clipAccountNumber.y1, (clipAccountNumber.x2 - clipAccountNumber.x1), (clipAccountNumber.y2 - clipAccountNumber.y1), page, cropDoc)
            CreateCropBox("NAME & ADDRESS", clipNA.x1, clipNA.y1, (clipNA.x2 - clipNA.x1), (clipNA.y2 - clipNA.y1), page, cropDoc)
            CreateCropBox("QESP STRING", clipQESP.x1, clipQESP.y1, (clipQESP.x2 - clipQESP.x1), (clipQESP.y2 - clipQESP.y1), page, cropDoc)

            cropDoc.Save(FRMW.getParameter("WORKDIR") & "\crop.pdf", SDF.SDFDoc.SaveOptions.e_compatibility + SDF.SDFDoc.SaveOptions.e_remove_unused)
        End Using

    End Sub

    Private Sub CreateCropBox(ByVal labelValue As String, ByVal x1Val As Double, ByVal y1Val As Double, ByVal x2Val As Double, ByVal y2Val As Double, ByVal PDFpage As Page, cropDoc As PDFDoc, Optional color1 As Double = 0.75, Optional color2 As Double = 0.75, Optional color3 As Double = 0.75, Optional opac As Double = 0.5)

        Dim elmBuilder As New ElementBuilder
        Dim elmWriter As New ElementWriter
        Dim element As Element
        elmWriter.Begin(PDFpage)
        elmBuilder.Reset() : elmBuilder.PathBegin()

        'Set crop box
        elmBuilder.CreateRect(x1Val, y1Val, x2Val, y2Val)
        elmBuilder.ClosePath()

        element = elmBuilder.PathEnd()
        element.SetPathFill(True)

        Dim gState As GState = element.GetGState
        gState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        gState.SetFillColor(New ColorPt(color1, color2, color3)) 'Default is gray
        gState.SetFillOpacity(opac)
        elmWriter.WriteElement(element)

        'Set text
        element = elmBuilder.CreateTextBegin(PDF.Font.Create(cropDoc, PDF.Font.StandardType1Font.e_helvetica_oblique, True), 8)
        element.GetGState.SetTextRenderMode(GState.TextRenderingMode.e_fill_text)
        element.GetGState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        element.GetGState.SetFillColor(New ColorPt(0, 0, 0))
        elmWriter.WriteElement(element)
        element = elmBuilder.CreateTextRun(labelValue)
        element.SetTextMatrix(1, 0, 0, 1, x1Val, (y1Val - 8))
        elmWriter.WriteElement(element)
        elmWriter.WriteElement(elmBuilder.CreateTextEnd())

        elmWriter.End()

    End Sub

    Private Sub LoadFonts(doc As PDFDoc)
        HelveticaRegularFont = pdftron.PDF.Font.Create(doc, PDF.Font.StandardType1Font.e_helvetica, False)
    End Sub

    Private Function GetPDFpageValue(clipRect As Rect) As String

        Dim docXML As New XmlDocument
        Dim X, Y, prevY As Double
        Dim x1Content As Double = clipRect.x1
        Dim y1Content As Double = clipRect.y1
        Dim x2Content As Double = clipRect.x2
        Dim y2Content As Double = clipRect.y2
        Dim contentValue As String = ""

        Using txt As TextExtractor = New TextExtractor
            Dim txtXML As String
            txt.Begin(CurrentPage, clipRect)
            txtXML = txt.GetAsXML(TextExtractor.XMLOutputFlags.e_output_bbox)
            docXML.LoadXml(txtXML)

            Dim tempRoot As XmlElement = docXML.DocumentElement
            Dim tempxnl1 As XmlNodeList
            tempxnl1 = Nothing
            tempxnl1 = tempRoot.SelectNodes("Flow/Para/Line")
            prevY = 0
            For Each elmC As XmlElement In tempxnl1
                Dim pos() As String = elmC.GetAttribute("box").Split(","c)
                X = pos(0) : Y = pos(1)

                'Page(Content)
                If (X >= x1Content) And (Y >= y1Content) And (X <= x2Content) And (Y <= y2Content) Then
                    If contentValue = "" Then
                        If prevY <> Math.Round(Y, 3) Then
                            contentValue = elmC.InnerText.Replace(vbLf, "")
                        End If
                    Else
                        contentValue = contentValue & elmC.InnerText.Replace(vbLf, "")
                    End If
                End If

                prevY = Math.Round(Y, 3)
                elmC = Nothing
            Next
        End Using

        Return contentValue

    End Function

    Private Function GetPDFpageValues(clipRect As Rect) As StringCollection
        Dim docXML As New XmlDocument
        Dim X, Y, prevY As Double
        Dim x1Content As Double = clipRect.x1
        Dim y1Content As Double = clipRect.y1
        Dim x2Content As Double = clipRect.x2
        Dim y2Content As Double = clipRect.y2
        Dim Values As New StringCollection

        Using txt As TextExtractor = New TextExtractor
            Dim txtXML As String
            txt.Begin(CurrentPage, clipRect)
            txtXML = txt.GetAsXML(TextExtractor.XMLOutputFlags.e_output_bbox)
            docXML.LoadXml(txtXML)

            Dim tempRoot As XmlElement = docXML.DocumentElement
            Dim tempxnl1 As XmlNodeList
            tempxnl1 = Nothing
            tempxnl1 = tempRoot.SelectNodes("Flow/Para/Line")
            prevY = 0
            For Each elmC As XmlElement In tempxnl1
                Dim pos() As String = elmC.GetAttribute("box").Split(","c)
                X = pos(0) : Y = pos(1)

                If (X >= x1Content) And (Y >= y1Content) And (X <= x2Content) And (Y <= y2Content) Then
                    If prevY <> Math.Round(Y, 3) Then
                        Values.Add(elmC.InnerText.Replace(vbLf, ""))
                    Else
                        Values(Values.Count - 1) = Values(Values.Count - 1) & elmC.InnerText.Replace(vbLf, "")
                    End If
                End If

                prevY = Math.Round(Y, 3)
                elmC = Nothing
            Next
        End Using

        Return Values
    End Function

    Private Function I2P(i As Decimal) As Decimal
        Return (i * 72)
    End Function

    Private Function CollectionToArray(Collection As StringCollection, ArraySize As Integer) As String()
        Dim Values(ArraySize) As String
        Dim i As Integer = 0
        For i = 0 To ArraySize
            If i <= Collection.Count - 1 Then
                Values(i) = Collection(i)
            Else
                Values(i) = ""
            End If
        Next
        Return Values
    End Function

    Private Sub WriteOutText(page As Page, textToWrite As String, xPosition As Double, yPosition As Double, Optional fontType As String = "REGULAR", Optional fontSize As Double = 10)
        Dim eb As New ElementBuilder
        Dim writer As New ElementWriter
        Dim element As Element
        writer.Begin(page)
        eb.Reset() : eb.PathBegin()

        element = eb.CreateTextBegin()
        element.GetGState.SetTextRenderMode(GState.TextRenderingMode.e_fill_text)
        element.GetGState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        element.GetGState.SetFillColor(New ColorPt(0, 0, 0))
        writer.WriteElement(element)

        Select Case fontType.ToUpper
            Case "REGULAR"
                'Helvetica
                element = eb.CreateTextRun(textToWrite, HelveticaRegularFont, fontSize)
            Case Else
                Throw New Exception(convLog.addError("Incorrect font type used in code, have tech take a look.", fontType.ToUpper, "REGULAR or BOLD", , , , , True))
        End Select

        Dim textWidth As Decimal = element.GetTextLength
        xPosition = I2P(xPosition)
        element.SetTextMatrix(1, 0, 0, 1, xPosition, I2P(yPosition))
        writer.WriteElement(element)
        writer.WriteElement(eb.CreateTextEnd())
        writer.End()

    End Sub

    Private Sub WhiteOutContentBox(x1Val As Double, y1Val As Double, x2Val As Double, y2Val As Double, Optional color1 As Double = 255, Optional color2 As Double = 255, Optional color3 As Double = 255, Optional opac As Double = 0.5)

        Dim elmBuilder As New ElementBuilder
        Dim elmWriter As New ElementWriter
        Dim element As Element
        elmWriter.Begin(CurrentPage)
        elmBuilder.Reset() : elmBuilder.PathBegin()

        'Set crop box
        elmBuilder.CreateRect(I2P(x1Val), I2P(y1Val), I2P(x2Val), I2P(y2Val))
        elmBuilder.ClosePath()

        element = elmBuilder.PathEnd()
        element.SetPathFill(True)

        Dim gState As GState = element.GetGState
        gState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        gState.SetFillColor(New ColorPt(color1, color2, color3)) 'default color is white
        gState.SetFillOpacity(opac)
        elmWriter.WriteElement(element)

        elmWriter.End()

    End Sub

    Private Sub AdjustPagePosition(PDFpage As Page, xPosition As Double, yPosition As Double)
        Dim element As Element
        Dim EW As ElementWriter
        Dim builder As ElementBuilder = New ElementBuilder
        'element = builder.CreateForm(PDFpage)

        'Dim currBox As Page.Box
        'currBox.
        Dim areaBox As New Rect
        areaBox = PDFpage.GetMediaBox
        areaBox.x1 = I2P(2)
        areaBox.Update()

        'areaBox.x1 = 0.7
        'areaBox.y1 = 2.6
        'areaBox.x2 = 2.25
        'areaBox.y2 = 0.9
        'element = builder.CreateRect(0.7, 2.6, 2.25, 0.9)
        ''element.GetBBox(areaBox)
        ''PDFpage.GetBox(areaBox)
        ''PDFpage.SetBox()
        'EW = New ElementWriter
        'EW.Begin(PDFpage, ElementWriter.WriteMode.e_replacement)
        'element.GetGState().SetTransform(1, 0, 0, 1, I2P(xPosition), I2P(-0.5))
        'EW.WritePlacedElement(element)
        'EW.End()
    End Sub

#End Region

#Region "Global Functions/Routines"

    Private Sub writeDOCUrecord(totalPages As Integer)
        DOCU.Clear()
        DOCU.AccountNumber = accountNumber
        DOCU.DocumentID = documentID
        DOCU.ClientCode = clientCode
        DOCU.DocumentDate = FRMW.getParameter("MM/DD/YYYY")
        DOCU.DocumentType = "Renewal"
        DOCU.AmountDue = "0"
        DOCU.DocumentKey = ""
        DOCU.Print_StartPage = StartingPage
        DOCU.Print_NumberOfPages = totalPages

        'Name/Address Info
        nameAddressList(0) = "" 'Set mailing ID to blank
        removeLastAddressLine(nameAddressList) 'Remove IMB data
        If foreignAddress Then
            DOCU.Print_HandlingCode = "F"
            DOCU.OriginalAddressForeign = True
        Else
            DOCU.Print_HandlingCode = "M"
        End If
        DOCU.setOriginalAddress(CollectionToArray(nameAddressList, 5), 1, foreignAddress)

        swDOCU.WriteLine(DOCU.GetXML)

    End Sub

    Private Sub removeLastAddressLine(addressList As StringCollection)
        Dim addressLine As String = addressList(addressList.Count - 1)
        addressLine = addressLine.Replace("A", "").Replace("D", "").Replace("F", "").Replace("T", "")
        If addressLine.Trim = "" Then
            addressList(addressList.Count - 1) = ""
        End If
    End Sub

#End Region

End Class
