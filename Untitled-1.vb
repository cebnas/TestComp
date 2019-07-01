Imports Infragistics.Win.UltraWinGrid
Imports System.Data.SqlClient
Imports System.Reflection.MethodBase
Imports Newtonsoft.Json
Imports Infragistics.Win
Imports System.IO

Public Class frmDeliverySheet
    Public RepNo As Integer
    Dim strSFormula As String
    Dim strRepPath As String
    Public bFromScanBarcode As Boolean = False
    Public iLineDetailID As Integer = 0
    Public iOrderIndex As Integer = 0
    Public pubMeIsNewRecord As Boolean
    Public pubMeMode As Integer = 0     '1 = New, 2 = Edit, 3 = Confirm Delivery
    Public thisTimeDel As Integer = 0
    Dim ugSOActiveRow As UltraGridRow
    Dim oProdDef As clsProdDefaults
    Dim oSoDefaults As New clsSOModuleDefaults
    Dim lastAccountId As Integer
    Dim ScanType As Integer = 0

    Sub PrintReport(ByVal sFormula As String, ByVal repPath As String)
        frmReportViewer.vSelectionFormula = sFormula
        frmReportViewer.vReportPath = repPath
        frmReportViewer.ShowDialog()
    End Sub

    Private Sub GET_AREAS() '
        Dim DS_BATCHES As DataSet
        SQL = "SELECT  idAreas, Description FROM Areas order by Description"
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                DS_BATCHES = .GET_INSERT_UPDATE(SQL)
                cboArea.DataSource = DS_BATCHES.Tables(0)
                cboArea.DisplayMember = "Description"
                cboArea.ValueMember = "idAreas"
                cboArea.DisplayLayout.Bands(0).Columns(0).Hidden = True
                cboArea.DisplayLayout.Bands(0).Columns(1).Width = cboArea.Width
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    Private Sub GetNextRunnID()
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                strSQL = "select Max(RunnNo) as MaxID from spilRunnSheetHeader"
                txtRunnNumber.Value = .GET_NEXT_NUMBER(strSQL)
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                objSQL = Nothing
            End Try
        End With
    End Sub

    Public Sub getDrivers()
        SQL = "SELECT     ID, Name, driverEmail FROM spilDriverMaster"
        Dim oSQL As New clsSqlConn
        With oSQL
            DS = New DataSet()
            DS = .GET_DATA_SQL(SQL)
            txtDrivName.DataSource = DS.Tables(0)
            txtDrivName.ValueMember = "driverEmail"
            txtDrivName.DisplayMember = "Name"
        End With
    End Sub

    Public Sub getVehicle()
        SQL = "SELECT     ID, Name FROM spilVehicleMaster"
        'SQL = "SELECT     ID, Name FROM spilVehicleMaster"
        Dim oSQL As New clsSqlConn
        With oSQL
            DS = New DataSet()
            DS = .GET_DATA_SQL(SQL)
            txtVehRegNo.DataSource = DS.Tables(0)
            txtVehRegNo.ValueMember = "ID"
            txtVehRegNo.DisplayMember = "Name"
        End With
    End Sub

    Private Sub frmDeliverySheet_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        DeleteUnnecessaryTrolleyData()
        DeleteUnnecessaryScannedData()
        Me.Dispose()
    End Sub

    Public Sub DeleteUnnecessaryTrolleyData()
        Dim oSQL As New clsSqlConn

        strSQL = "DELETE FROM spil_DespatchScannedTrolleyBarcodes WHERE RunnNo = -999"
        oSQL.Exe_Query(strSQL)
    End Sub

    Public Sub DeleteUnnecessaryScannedData()
        Dim oSQL As New clsSqlConn

        strSQL = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE RunnNo = -999"
        oSQL.Exe_Query(strSQL)
    End Sub

    Private Sub frmDeliverySheet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call GET_AREAS()
        Call GET_Facility()
        Call getDrivers()
        Call getVehicle()
        Call SetGridProperties()

        If pubMeMode = 1 Then 'New
            txtRunningDate.Value = Now.Date
            txtRunnTime.Value = CDate(Now.Date & " 08:00:00 AM")
            cmbFacility.Value = 1

            tsbConfirmDeliv.Visible = False
            ConfirmDeliveryAllToolStripMenuItem.Visible = False
            tsbSave.Visible = True
            tsbScanOptions.Visible = True

        ElseIf pubMeMode = 2 Then 'Edit
            tsbConfirmDeliv.Visible = False
            ConfirmDeliveryAllToolStripMenuItem.Visible = False
            tsbSave.Visible = True
            tsbScanOptions.Visible = False

        ElseIf pubMeMode = 3 Then 'DeliveyConfirm
            tsbConfirmDeliv.Visible = True
            ConfirmDeliveryAllToolStripMenuItem.Visible = True
            tsbSave.Visible = False
            tsbScanOptions.Visible = False
        End If

        Call VisibleSelectByPieceBCode()

        cboArea.Focus()

        Dim objDespatchDef As New clsDespatchDefaults
        'If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectOrders Or objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
        If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
            tsmRefreshQuantities.Visible = False
        End If
        GoogleMapControllerHandler(isGoogleAPIActive)
    End Sub

    Private Sub VisibleSelectByPieceBCode()
        Dim objDespatchDef As New clsDespatchDefaults

        'If IsPieceTrackingEnabled = True Then
        '    If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
        '        tsbSelectByPieceBCode.Visible = True
        '    Else
        '        tsbSelectByPieceBCode.Visible = False
        '    End If
        'Else
        '    tsbSelectByPieceBCode.Visible = False
        'End If

        If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
            'UGSOList.DisplayLayout.Bands(0).Columns("TotalGlassPanels").Hidden = True
            tsbSelectBCode.Visible = True
            tsbSelFromList.Visible = True
            tsbSelectByPieceBCode.Visible = True
        ElseIf objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
            'UGSOList.DisplayLayout.Bands(0).Columns("TotalGlassPanels").Hidden = True
            tsbSelectBCode.Visible = True
            tsbSelFromList.Visible = True
            tsbSelectByPieceBCode.Visible = False
        End If
    End Sub

    Private Sub GET_Facility()
        SQL = "SELECT     FacilityID, FacilityName  FROM spilPROD_Facility ORDER BY FacilityID"   '
        Dim objSQL As New clsSqlConn
        With objSQL
            Try
                Dim DS_BATCHES As DataSet = .GET_INSERT_UPDATE(SQL)
                cmbFacility.DataSource = DS_BATCHES.Tables(0)
                cmbFacility.DisplayMember = "FacilityName"
                cmbFacility.ValueMember = "FacilityID"
                cmbFacility.DisplayLayout.Bands(0).Columns(0).Hidden = True ' Width = 139
                cmbFacility.DisplayLayout.Bands(0).Columns(1).Width = cmbFacility.Width '200

                If cmbFacility.Rows.Count > 0 Then
                    cmbFacility.Value = cmbFacility.GetRow(ChildRow.First).Cells("FacilityID").Value
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
            Finally
                .Dispose()
                objSQL = Nothing
            End Try
        End With
    End Sub

    Private Sub SetGridProperties()
        'Grid properties for UGSOLIST
        UGSOList.DisplayLayout.Bands(0).Columns("DocType").Hidden = True
        UGSOList.DisplayLayout.Bands(0).Columns("DocState").Hidden = True
        UGSOList.DisplayLayout.Bands(0).Columns("iAreasID").Hidden = True
        UGSOList.DisplayLayout.Bands(0).Columns("RunnStatus").Hidden = True

        'If My.Computer.FileSystem.FileExists(strAppPath & "\" & strUserName & "RunningSheetDetails.xml") Then
        '    UGSOList.DisplayLayout.LoadFromXml(strAppPath & "\" & strUserName & "RunningSheetDetails.xml")
        'End If


        UGSOList.DisplayLayout.Bands(0).Columns("OrderFinishedItems").Header.Caption = "Ordered Items"
        UGSOList.DisplayLayout.Bands(0).Columns("TotalGlassPanels").Header.Caption = "Total Glass Panels"

        UGSOList.DisplayLayout.Bands(0).Override.HeaderAppearance.TextHAlign = HAlign.Center
        UGSOList.DisplayLayout.Bands(0).Override.HeaderAppearance.TextVAlign = VAlign.Middle
        UGSOList.DisplayLayout.Bands(0).Override.FilterRowAppearance.BackColor = Color.Beige
        UGSOList.DisplayLayout.Bands(0).Override.RowAlternateAppearance.BackColor = Color.FromArgb(236, 242, 253)
        UGSOList.DisplayLayout.Bands(0).Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Solid
        UGSOList.DisplayLayout.Bands(0).Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.None
        UGSOList.DisplayLayout.Bands(0).Override.HeaderAppearance.BackColor = Color.SteelBlue
        UGSOList.DisplayLayout.Bands(0).Override.HeaderAppearance.ForeColor = Color.White
        UGSOList.DisplayLayout.Bands(0).Override.HeaderAppearance.BackGradientStyle = GradientStyle.None
        UGSOList.DisplayLayout.Bands(0).Override.CellAppearance.TextVAlign = VAlign.Middle
        UGSOList.DisplayLayout.Bands(0).Override.CellClickAction = CellClickAction.RowSelect
        UGSOList.DisplayLayout.Bands(0).Override.SelectTypeRow = SelectType.Extended

        UGSOList.Update()

        'Added by Hashini on 18-10-2018 - Grid properties for UGSOLines
        If My.Computer.FileSystem.FileExists(strAppPath & "\" & strUserName & "RunningSheetSODetails.xml") Then
            UGSOLines.DisplayLayout.LoadFromXml(strAppPath & "\" & strUserName & "RunningSheetSODetails.xml")
        End If

        UGSOLines.DisplayLayout.Bands(0).Override.HeaderAppearance.TextHAlign = HAlign.Center
        UGSOLines.DisplayLayout.Bands(0).Override.HeaderAppearance.TextVAlign = VAlign.Middle
        UGSOLines.DisplayLayout.Bands(0).Override.FilterRowAppearance.BackColor = Color.Beige
        UGSOLines.DisplayLayout.Bands(0).Override.RowAlternateAppearance.BackColor = Color.FromArgb(236, 242, 253)
        UGSOLines.DisplayLayout.Bands(0).Override.BorderStyleCell = UIElementBorderStyle.Solid
        UGSOLines.DisplayLayout.Bands(0).Override.BorderStyleRow = UIElementBorderStyle.None
        UGSOLines.DisplayLayout.Bands(0).Override.HeaderAppearance.BackColor = Color.SteelBlue
        UGSOLines.DisplayLayout.Bands(0).Override.HeaderAppearance.ForeColor = Color.White
        UGSOLines.DisplayLayout.Bands(0).Override.HeaderAppearance.BackGradientStyle = GradientStyle.None
        UGSOLines.DisplayLayout.Bands(0).Override.CellAppearance.TextVAlign = VAlign.Middle
        UGSOLines.DisplayLayout.Bands(0).Override.CellAppearance.TextHAlign = HAlign.Center
        UGSOLines.DisplayLayout.Bands(0).Override.CellClickAction = CellClickAction.RowSelect
        UGSOLines.DisplayLayout.Bands(0).Override.SelectTypeRow = SelectType.Extended

        UGSOLines.Update()
    End Sub


    Public Sub GetRunningSheetData()

        Dim objSQL As New clsSqlConn
        Dim DS_ITEMS As DataSet
        Dim dr1 As DataRow
        Dim ugR As UltraGridRow
        Dim objDespatchDef As New clsDespatchDefaults
        Dim iOrderQuantity As Integer = 0

        Try

            SQL = "select * from spilRunnSheetHeader where RunnNo=" & txtRunnNumber.Value & " "
            SQL += " SELECT spilRunnSheetDetail.RecID, spilRunnSheetDetail.ThisTimeQty, " &
                "spilRunnSheetDetail.OrderIndex, spilRunnSheetDetail.Status,spilRunnSheetDetail.Comment, spilPROD_STATES.ProdStateName, spilInvNum.OrderNum, " &
                "spilInvNum.Delivery_Status, Client.DCLink, Client.Name, spilInvNum.OrderDate, " &
                "spilInvNum.ExtOrderNum, spilInvNum.TotalFinishedItems AS OrderFinishedItems, " &
                "spilInvNum.TotalGlassPanels AS TotalGlassPanels, " &
                "spilInvNum.DeliveredFinishedItems AS DeliveredFinishedItems, Areas.idAreas, " &
                "Areas.Description, spilInvNum.DueDate,spilPROD_STATES.ProductionState, spilInvNum.DocType, spilRunnSheetDetail.geoCoordinations " &
                "FROM spilRunnSheetDetail INNER JOIN " &
                "spilInvNum ON spilRunnSheetDetail.OrderIndex = spilInvNum.OrderIndex INNER JOIN " &
                "spilPROD_STATES ON spilInvNum.ProductionState = spilPROD_STATES.ProductionState INNER JOIN " &
                "Client ON spilInvNum.AccountID = Client.DCLink LEFT OUTER JOIN " &
                "Areas ON spilInvNum.iAreasID = Areas.idAreas where RunnNo=" & txtRunnNumber.Value & " "
            SQL += " SELECT * FROM spilRunnSheetDetailLines WHERE RunnNo= " & txtRunnNumber.Value & " "

            DS_ITEMS = objSQL.GET_INSERT_UPDATE(SQL)

            For Each dr1 In DS_ITEMS.Tables(1).Rows

                ugR = UGSOList.DisplayLayout.Bands(0).AddNew

                ugR.Cells("RecID").Value = dr1("RecID")
                ugR.Cells("OrderIndex").Value = dr1("OrderIndex")
                ugR.Cells("DocState").Value = dr1("ProductionState")
                ugR.Cells("RunnStatus").Value = dr1("Status")
                ugR.Cells("OrdStateName").Value = dr1("ProdStateName")
                ugR.Cells("DelState").Value = dr1("Delivery_Status")
                ugR.Cells("OrderNum").Value = dr1("OrderNum")
                ugR.Cells("AccountID").Value = dr1("DCLink")
                ugR.Cells("CustomerName").Value = dr1("Name")
                ugR.Cells("OrderDate").Value = dr1("OrderDate")
                ugR.Cells("ExtOrderNum").Value = dr1("ExtOrderNum")
                ugR.Cells("iAreasID").Value = dr1("idAreas")
                ugR.Cells("AreaName").Value = dr1("Description")
                ugR.Cells("DueDate").Value = dr1("DueDate")
                ugR.Cells("DocType").Value = dr1("DocType")
                ugR.Cells("OrderFinishedItems").Value = dr1("OrderFinishedItems")
                'ugR.Cells("DeliveredSoFar").Value = dr1("DeliveredFinishedItems")
                ugR.Cells("ThisTimeDelivery").Value = dr1("ThisTimeQty")
                ugR.Cells("TotalGlassPanels").Value = dr1("TotalGlassPanels")
                ugR.Cells("Comment").Value = dr1("Comment")
                ugR.Cells("geoCoordinations").Value = dr1("geoCoordinations")

                If ugR.Cells("RunnStatus").Value = DeliveryState.Delivered Then
                    ugR.Activation = Activation.Disabled
                End If

                iDeliveredAndPendingQty = FillDeliveredSoFarQty(dr1("OrderIndex"), ugR.Cells("RunnStatus").Value)

                If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
                    iOrderQuantity = ugR.Cells("TotalGlassPanels").Value
                Else
                    iOrderQuantity = ugR.Cells("OrderFinishedItems").Value
                End If

                If iOrderQuantity >= iDeliveredAndPendingQty Then
                    ugR.Cells("DeliveredSoFar").Value = iDeliveredAndPendingQty
                Else
                    ugR.Cells("DeliveredSoFar").Value = iOrderQuantity
                End If

                If (ugR.Cells("ThisTimeDelivery").Value < iOrderQuantity) Then
                    ugR.Cells("ThisTimeDelivery").Appearance.BackColor = Color.Yellow
                    ugR.Cells("ThisTimeDelivery").Appearance.ForeColor = Color.Red
                Else
                    ugR.Cells("ThisTimeDelivery").Value = iOrderQuantity
                End If

                If (CInt(ugR.Cells("ThisTimeDelivery").Value) + CInt(ugR.Cells("DeliveredSoFar").Value)) < iOrderQuantity Then
                    ugR.Appearance.ForeColor = Color.FromArgb(130, 7, 7)
                Else
                    ugR.Appearance.ForeColor = Color.FromArgb(4, 73, 6)
                End If

                If (ugR.Cells("DocType").Value = GlassDocTypes.NCR) Then
                    ugR.Appearance.ForeColor = Color.SaddleBrown
                End If
            Next

            For Each dr1 In DS_ITEMS.Tables(0).Rows
                If dr1("Status") <> GlassReceiptState.Processed Then
                    tsbSave.Enabled = True
                    tsbPrint.Enabled = True
                Else
                    tsbSave.Enabled = False
                    tsbPrint.Enabled = True
                    tsmRefreshQuantities.Visible = False
                End If

                cboArea.Value = dr1("AreaID")
                txtReference.Text = dr1("Reference")
                txtDrivName.Text = dr1("DrivName")
                txtVehRegNo.Text = dr1("VehRegNo")
                txtTeleNo.Text = dr1("TelNo")
                txtNotes.Text = dr1("Notes")
                txtRunningDate.Value = dr1("RunnDate")
                txtRunnTime.Value = dr1("RunnTime")
                cmbFacility.Value = dr1("FacilityID")
                cmbDuration.Value = dr1("Duration")
                If dr1("googleMapImagePath") <> "" AndAlso isGoogleAPIActive = True Then
                    pbGoogleMap.Image = Image.FromFile(dr1("googleMapImagePath"))
                    btnRefreshGoogleMap.Visible = False
                End If
            Next

            For Each drSOLine As DataRow In DS_ITEMS.Tables(2).Rows
                ugR = UGSOLines.DisplayLayout.Bands(0).AddNew()
                ugR.Cells("OrderIndex").Value = drSOLine("OrderIndex")
                ugR.Cells("iInvDetailID").Value = drSOLine("iInvDetailID")
                ugR.Cells("LineNo").Value = drSOLine("LineNumber")
                ugR.Cells("Description").Value = drSOLine("Description")
                ugR.Cells("Thickness").Value = drSOLine("Thickness")
                ugR.Cells("Height").Value = drSOLine("Height")
                ugR.Cells("Width").Value = drSOLine("Width")
                ugR.Cells("Size").Value = drSOLine("Height") & " X " & drSOLine("Width")
                ugR.Cells("OrderQty").Value = drSOLine("OrderQty")
                ugR.Cells("PrevDelQty").Value = drSOLine("PrevDelQty")
                ugR.Cells("RecutQty").Value = drSOLine("RecutQty")

                ugR.Cells("ThisDelQty").Activation = Activation.AllowEdit
                ugR.Cells("ThisDelQty").Value = drSOLine("ThisDelQty")
                ugR.Cells("GlassWeight").Value = drSOLine("GlassWeight")

                If drSOLine("LineTypeID") = LineState.NCR Then
                    ugR.Cells("LineType").Value = "NCR"
                ElseIf drSOLine("LineTypeID") = LineState.ReBatched Then
                    ugR.Cells("LineType").Value = "ReBatched"
                ElseIf drSOLine("LineTypeID") = LineState.Normal Then
                    ugR.Cells("LineType").Value = "Normal"
                End If

                ugR.Cells("LineTypeID").Value = drSOLine("LineTypeID")

                If (ugR.Cells("ThisDelQty").Value + ugR.Cells("PrevDelQty").Value + ugR.Cells("RecutQty").Value) < ugR.Cells("OrderQty").Value Then
                    ugR.CellAppearance.ForeColor = Color.FromArgb(130, 7, 7)
                    ugR.Cells("ThisDelQty").Appearance.BackColor = Color.Yellow
                Else
                    ugR.CellAppearance.ForeColor = Color.Green
                End If

                If drSOLine("LineTypeID") = LineState.ReBatched Then
                    ugR.CellAppearance.ForeColor = Color.Red
                End If
            Next

            UGSOLines.ActiveRow = Nothing

            Dim ugSORow As UltraGridRow = UGSOList.ActiveRow
            For Each ugRow As UltraGridRow In UGSOLines.Rows
                If ugRow.Cells("OrderIndex").Value = ugSORow.Cells("OrderIndex").Value Then
                    ugRow.Hidden = False
                Else
                    ugRow.Hidden = True
                End If
            Next

            GetVehicleDetails(False, False, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        Finally
            dr1 = Nothing
            DS_ITEMS = Nothing
            objSQL = Nothing
        End Try
    End Sub


    Private Sub tsbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbPrint.Click
        ''   strSFormula = "{spilInvNum.DueDate} = DateTime(" & DatePart(DateInterval.Year, txtRunningDate.Value) & "," & DatePart(DateInterval.Month, txtRunningDate.Value) & "," & DatePart(DateInterval.Day, txtRunningDate.Value) & ") " & _
        ''" and {spilInvNum.iAreasID} =" & cboArea.Value & " and {spilInvNumLines.ItemTypeCategory} in ['G', 'M'] and {spilInvNum.DocType} in [4, 7] and {spilInvNum.ProductionState} <> 6 "

        GlassReportType = "RUNN"
        strSFormula = "{spilRunnSheetHeader.RunnNo}=" & txtRunnNumber.Value & ""
        strRepPath = EvoGlassReportPath & "\Reports\Production\Running Sheet.rpt"
        PrintReport(strSFormula, strRepPath)

    End Sub

    Private Sub tsbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Public Sub UpdateThisTimeDelivery(ByVal iOrderIndex As Integer)
        Dim objSQL As New clsSqlConn
        With objSQL
            If (UGSOList.Rows.Count > 0) Then
                Dim ugR As UltraGridRow
                For Each ugR In UGSOList.Rows
                    If ugR.Cells("OrderIndex").Value = iOrderIndex Then

                        'If pubMeIsNewRecord = True Then
                        Dim sql As String = "SELECT Count(OrderIndex) FROM spilRunnSheetDownloadedBarCodes WHERE OrderIndex = " & iOrderIndex & " AND RunnNo=-999  AND Status='OK' GROUP BY OrderIndex"
                        Dim count As Integer = .Get_ScalerINTEGER(sql)

                        ugR.Cells("ThisTimeDelivery").Value = 0
                        ugR.Cells("ThisTimeDelivery").Value += count
                        'Else
                        '    ugR.Cells("ThisTimeDelivery").Value += 1
                        'End If

                        If (CInt(ugR.Cells("ThisTimeDelivery").Value) + CInt(ugR.Cells("DeliveredSoFar").Value) > CInt(ugR.Cells("OrderFinishedItems").Value)) Then
                            ugR.Cells("ThisTimeDelivery").Value = CInt(ugR.Cells("OrderFinishedItems").Value) - CInt(ugR.Cells("DeliveredSoFar").Value)
                        ElseIf (CInt(ugR.Cells("ThisTimeDelivery").Value) + CInt(ugR.Cells("DeliveredSoFar").Value) < CInt(ugR.Cells("OrderFinishedItems").Value)) Then
                            ugR.Cells("ThisTimeDelivery").Appearance.BackColor = Color.Yellow
                            ugR.Cells("ThisTimeDelivery").Appearance.ForeColor = Color.Red
                        Else
                            ugR.Cells("ThisTimeDelivery").Appearance.BackColor = Color.White
                            ugR.Cells("ThisTimeDelivery").Appearance.ForeColor = Color.Black
                        End If
                    End If
                Next
            End If


        End With
    End Sub

    'Public Sub AddOrderData(ByVal iOrderIndex As Integer)
    '    Dim objDespatchDef As New clsDespatchDefaults
    '    Dim objSQL As New clsSqlConn
    '    Dim DS_ITEMS As DataSet
    '    Dim dr1 As DataRow
    '    Dim ugR As UltraGridRow
    '    Dim strWhere As String
    '    Dim iThisTimeQty As Integer = 0
    '    Dim iDeliveredAndPendingQty As Integer = 0
    '    Dim iOrderQuantity As Integer = 0
    '    Dim booIsFound As Boolean = False

    '    For Each ugR In UGSOList.Rows
    '        If ugR.Cells("OrderIndex").Value = iOrderIndex Then
    '            booIsFound = True
    '            Exit Sub
    '        End If
    '    Next

    '    If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
    '        iThisTimeQty = GetDespathScheduledQuantity(iOrderIndex)
    '        If iThisTimeQty <= 0 Then Exit Sub
    '    ElseIf objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
    '        iThisTimeQty = GetManuallyScanPieces(iOrderIndex)
    '        If iThisTimeQty <= 0 Then Exit Sub
    '    End If


    '    SQL = "SELECT RunnNo, (SELECT OrderNum FROM spilInvNum WHERE OrderIndex = " & iOrderIndex & ") AS OrderNum FROM spilRunnSheetDetail WHERE OrderIndex = " & iOrderIndex & ""
    '    Dim dsPrevRuunNos As DataSet = objSQL.GET_DataSet(SQL)

    '    If dsPrevRuunNos.Tables(0).Rows.Count > 0 Then
    '        If MsgBox("This order (" & dsPrevRuunNos.Tables(0).Rows(0)("OrderNum") & ") has already been added to running sheet no: " & dsPrevRuunNos.Tables(0).Rows(0)("RunnNo") & ". Do you want to continue?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "SPIL Glass") = MsgBoxResult.No Then
    '            Exit Sub
    '        End If
    '    End If

    '    Try

    '        strWhere = " Where OrderIndex=" & iOrderIndex & " "

    '        SQL = "SELECT spilInvNum.OrderIndex, spilInvNum.DocType, spilInvNum.AccountID, " & _
    '            "spilInvNum.TotalFinishedItems AS OrderFinishedItems, " & _
    '            "spilInvNum.TotalGlassPanels AS TotalGlassPanels, " & _
    '            "spilInvNum.DeliveredFinishedItems AS DeliveredFinishedItems, " & _
    '            "Client.Account AS CustCode, spilInvNum.OrderNum, CONVERT(datetime, " & _
    '            "CONVERT(CHAR(10), spilInvNum.OrderDate, 103), 103) AS OrderDate, " & _
    '            "Client.Name AS Customer, spilInvNum.ExtOrderNum AS CustOrdNo, spilInvNum.DueDate, " & _
    '            "spilPROD_STATES.ProdStateName AS ProdState, " & _
    '            "spilPROD_DeliveryState.DeliveryDescription AS DelState, " & _
    '            "spilPROD_InvoiceState.DocStateName AS DocuState, spilPROD_STATES.ProductionState, " & _
    '            "spilPROD_InvoiceState.DocState, spilInvNum.DocState AS InvDocState, " & _
    '            "Areas.Description AS Area, spilInvNum.geoCoordinations " & _
    '            "FROM spilPROD_DeliveryState WITH (NOLOCK) RIGHT OUTER JOIN spilInvNum WITH (NOLOCK) " & _
    '            "LEFT OUTER JOIN Areas WITH (NOLOCK) ON spilInvNum.iAreasID = Areas.idAreas LEFT OUTER JOIN " & _
    '            "spilPROD_STATES WITH (NOLOCK) ON spilInvNum.ProductionState = spilPROD_STATES.ProductionState " & _
    '            "ON spilPROD_DeliveryState.Delivery_Status = spilInvNum.Delivery_Status LEFT OUTER JOIN " & _
    '            "spilPROD_InvoiceState WITH (NOLOCK) ON spilInvNum.DocState = spilPROD_InvoiceState.DocState " & _
    '            "LEFT OUTER JOIN Client WITH (NOLOCK) ON spilInvNum.AccountID = Client.DCLink " & _
    '            "" & strWhere & " ORDER BY spilInvNum.OrderNum"

    '        DS_ITEMS = objSQL.GET_INSERT_UPDATE(SQL)

    '        For Each dr1 In DS_ITEMS.Tables(0).Rows

    '            ugR = UGSOList.DisplayLayout.Bands(0).AddNew

    '            ugR.Cells("RecID").Value = 0
    '            ugR.Cells("OrderIndex").Value = dr1("OrderIndex")
    '            ugR.Cells("DocState").Value = 0     'dr1("InvDocState")
    '            ugR.Cells("RunnStatus").Value = 0
    '            ugR.Cells("DelState").Value = 0
    '            ugR.Cells("OrdStateName").Value = dr1("ProdState")
    '            ugR.Cells("OrderNum").Value = dr1("OrderNum")
    '            ugR.Cells("AccountID").Value = dr1("AccountID")
    '            ugR.Cells("CustomerName").Value = dr1("Customer")
    '            ugR.Cells("OrderDate").Value = dr1("OrderDate")
    '            ugR.Cells("ExtOrderNum").Value = dr1("CustOrdNo")
    '            ugR.Cells("iAreasID").Value = 0     'dr1("CustOrdNo")
    '            ugR.Cells("AreaName").Value = dr1("Area")
    '            ugR.Cells("DueDate").Value = dr1("DueDate")
    '            ugR.Cells("DocType").Value = dr1("DocType")
    '            ugR.Cells("OrderFinishedItems").Value = dr1("OrderFinishedItems")
    '            ugR.Cells("TotalGlassPanels").Value = dr1("TotalGlassPanels")
    '            ugR.Cells("geoCoordinations").Value = dr1("geoCoordinations")
    '            If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
    '                iOrderQuantity = ugR.Cells("TotalGlassPanels").Value
    '            Else
    '                iOrderQuantity = ugR.Cells("OrderFinishedItems").Value
    '            End If

    '            iDeliveredAndPendingQty = FillDeliveredSoFarQty(iOrderIndex, ugR.Cells("RunnStatus").Value)
    '            If iOrderQuantity >= iDeliveredAndPendingQty Then
    '                ugR.Cells("DeliveredSoFar").Value = iDeliveredAndPendingQty
    '            Else
    '                ugR.Cells("DeliveredSoFar").Value = iOrderQuantity
    '            End If

    '            If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.DespatchTotalOrderQuantity Then
    '                ugR.Cells("ThisTimeDelivery").Value = ugR.Cells("OrderFinishedItems").Value - ugR.Cells("DeliveredSoFar").Value
    '            Else
    '                If iOrderQuantity >= CInt(ugR.Cells("DeliveredSoFar").Value) + iThisTimeQty Then
    '                    ugR.Cells("ThisTimeDelivery").Value = iThisTimeQty
    '                Else
    '                    ugR.Cells("ThisTimeDelivery").Value = iOrderQuantity - CInt(ugR.Cells("DeliveredSoFar").Value)
    '                End If
    '            End If

    '            If CInt(ugR.Cells("ThisTimeDelivery").Value) < iOrderQuantity Then
    '                ugR.Cells("ThisTimeDelivery").Appearance.BackColor = Color.Yellow
    '                ugR.Cells("ThisTimeDelivery").Appearance.ForeColor = Color.Red
    '            End If
    '        Next

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
    '    Finally
    '        dr1 = Nothing
    '        DS_ITEMS = Nothing
    '        objSQL = Nothing
    '    End Try

    'End Sub

    Private Function FillDeliveredSoFarQty(OrderIndex As Integer, Status As Integer) As Integer
        Dim objSQL As New clsSqlConn
        Dim DeliveredQty As Integer = 0
        Dim RunnSheetConfirmQty As Integer = 0
        Dim TotalQty As Integer = 0

        strQuery = "SELECT DeliveredFinishedItems FROM spilInvNum WITH (NOLOCK) WHERE OrderIndex = " & OrderIndex & ""
        DeliveredQty = objSQL.Get_ScalerINTEGER(strQuery)


        If Status = DeliveryState.Delivered Then
            strQuery = "SELECT ISNULL(SUM(ThisTimeQty),0) As TotRunnSheetQty FROM spilRunnSheetDetail " &
            "INNER JOIN spilRunnSheetHeader ON spilRunnSheetDetail.RunnNo = spilRunnSheetHeader.RunnNo " &
            "WHERE spilRunnSheetDetail.OrderIndex = " & OrderIndex & " AND " &
            "spilRunnSheetDetail.RunnNo = " & txtRunnNumber.Value & ""

            RunnSheetConfirmQty = objSQL.Get_ScalerINTEGER(strQuery)
        End If

        TotalQty = (DeliveredQty - RunnSheetConfirmQty)

        Return TotalQty
    End Function

    Private Function GetDespathScheduledQuantity(OrderIndex As Integer) As Integer
        Dim objSQL As New clsSqlConn
        Dim strQuery As String
        Dim sDesSchBarcodes As New ArrayList
        Dim iThisTimeQty As Integer = 0
        Dim DelStation As Integer = 0
        Dim dsDesSchQty As DataSet

        objSQL.Begin_Trans()

        strQuery = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE SerialBarcodeValue " &
            "IN (SELECT BarCodeV FROM spilPROD_SERIALS WHERE Qty_ReBatched = 1 AND " &
            "OrderIndex = " & OrderIndex & ") AND Status = 'AUTO'"
        If objSQL.Exe_Query_Trans(strQuery) = 0 Then
            objSQL.Rollback_Trans()
            Exit Function
        End If

        strQuery = "SELECT TOP(1) spilPROD_BATCH.STATION_TP_ID FROM spilInvNumLines WITH (NOLOCK) INNER JOIN " &
            "spilPROD_BATCH WITH (NOLOCK) ON spilInvNumLines.iInvDetailID = spilPROD_BATCH.iInvDetailID " &
            "WHERE (spilInvNumLines.OrderIndex = " & OrderIndex & ") ORDER BY spilPROD_BATCH.ProcessPath DESC"

        DelStation = objSQL.Get_ScalerINTEGER_WithTrans(strQuery)
        If DelStation <= 0 Then Exit Function

        strQuery = "SELECT BarCodeV FROM spilPROD_SERIALS WITH (NOLOCK) WHERE " &
            "OrderIndex = " & OrderIndex & " AND STATION_TP_ID = " & DelStation & " AND Qty_In > 0 " &
            "AND Qty_Out = 0 AND BarCodeV NOT IN (SELECT BarCodeV FROM spilPROD_SERIALS WHERE " &
            "Qty_ReBatched = 1 AND OrderIndex = " & OrderIndex & ");"
        strQuery += "SELECT RecID, RunnNo, SerialBarcodeValue, OrderIndex FROM spilRunnSheetDownloadedBarCodes " &
            "WHERE OrderIndex = " & OrderIndex & ""
        dsDesSchQty = objSQL.Get_Data_Trans(strQuery)

        For Each row As DataRow In dsDesSchQty.Tables(0).Rows
            sDesSchBarcodes.Add(row("BarCodeV"))
        Next

        For Each row As DataRow In dsDesSchQty.Tables(1).Rows
            If (sDesSchBarcodes.Contains(row("SerialBarcodeValue"))) Then
                sDesSchBarcodes.Remove(row("SerialBarcodeValue"))
            End If
        Next

        For Each sDesSchBarcode As String In sDesSchBarcodes
            strQuery = "SELECT iInvDetailID FROM spilPROD_SERIALS WHERE BarCodeV = '" & sDesSchBarcode & "'"
            Dim iInvDetailID As Integer = objSQL.Get_ScalerINTEGER_WithTrans(strQuery)

            strQuery = "SET DATEFORMAT DMY INSERT INTO spilRunnSheetDownloadedBarCodes " &
                "(RunnNo,BarcodeValue,BarcodeTrolley,TaggedTime,Status,Qty,SerialBarcodeValue,OrderIndex,iInvDetailID) " &
                "VALUES (-999,'','','" & Now & "','AUTO',1,'" & sDesSchBarcode & "'," & OrderIndex & "," & iInvDetailID & ")"
            If objSQL.Exe_Query_Trans(strQuery) = 0 Then
                objSQL.Rollback_Trans()
                Exit Function
            End If
        Next

        objSQL.Commit_Trans()

        iThisTimeQty = sDesSchBarcodes.Count

        Return iThisTimeQty
    End Function

    Private Function GetManuallyScanPieces(OrderIndex As Integer) As Integer
        Dim objSQL As New clsSqlConn
        Dim strQuery As String
        Dim sDesSchBarcodes As New ArrayList
        Dim iScanPieces As Integer = 0
        Dim dsScanPieces As DataSet

        objSQL.Begin_Trans()

        strQuery = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE SerialBarcodeValue " &
            "IN (SELECT BarCodeV FROM spilPROD_SERIALS WHERE Qty_ReBatched = 1 AND " &
            "OrderIndex = " & OrderIndex & ") AND Status = 'OK'"
        If objSQL.Exe_Query_Trans(strQuery) = 0 Then
            objSQL.Rollback_Trans()
            Exit Function
        End If

        strQuery = "SELECT COUNT(BarcodeValue) FROM spilRunnSheetDownloadedBarCodes " &
           "WHERE OrderIndex = " & OrderIndex & " AND Status = 'OK' AND RunnNo = -999"

        iScanPieces = objSQL.Get_ScalerINTEGER_WithTrans(strQuery)

        objSQL.Commit_Trans()

        Return iScanPieces
    End Function

    Private Sub SaveRunningSheet(Optional ByRef saveOnly As Boolean = False)
        Dim objSQL As New clsSqlConn
        Dim objDespatchDef As New clsDespatchDefaults

        With objSQL
            Try
                ''Dim bFoundMissingItems As Boolean = False

                If pubMeIsNewRecord = True Then
                    Call GetNextRunnID()
                End If

                .Begin_Trans()

                'If pubMeIsNewRecord = True Then
                '    strSQL = "set dateformat dmy Insert into spilRunnSheetHeader " &
                '        "(RunnNo, AreaID, RunnDate, RunnTime, VehRegNo, DrivName, Reference, TelNo, Notes, " &
                '        "DocPrinted, EnteredBy, EnteredDateTime, Status, FacilityID, Duration) " &
                '        "values(" & txtRunnNumber.Value & ", " & cboArea.Value & ", '" & txtRunningDate.Value & "', " &
                '        "'" & txtRunnTime.Value & "', '" & txtVehRegNo.Text.Replace("'", "") & "', " &
                '        "'" & txtDrivName.Text.Replace("'", "") & "', '" & txtReference.Text.Replace("'", "") & "', " &
                '        "'" & txtTeleNo.Text.Replace("'", "") & "', '" & txtNotes.Text.Replace("'", "") & "', 0, " &
                '        "'" & strUserName & "', '" & Now & "', " & GlassReceiptState.UnProcessed & ", " &
                '        "" & cmbFacility.Value & ", '" & cmbDuration.Value & "')"
                'Else
                '    strSQL = "set dateformat dmy update spilRunnSheetHeader " &
                '        "set AreaID=" & cboArea.Value & "," &
                '        "RunnDate='" & txtRunningDate.Value & "'," &
                '        "RunnTime='" & txtRunnTime.Value & "'," &
                '        "VehRegNo='" & txtVehRegNo.Text.Replace("'", "") & "'," &
                '        "DrivName='" & txtDrivName.Text.Replace("'", "") & "'," &
                '        "Reference='" & txtReference.Text.Replace("'", "") & "'," &
                '        "TelNo='" & txtTeleNo.Text.Replace("'", "") & "'," &
                '        "Notes='" & txtNotes.Text.Replace("'", "") & "'," &
                '        "EnteredBy='" & strUserName & "'," &
                '        "EnteredDateTime='" & Now & "', " &
                '        "FacilityID=" & cmbFacility.Value & ", " &
                '        "Duration='" & cmbDuration.Value & "' " &
                '        "where RunnNo=" & txtRunnNumber.Value & ""
                'End If

                'If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                '    .Rollback_Trans()
                '    Exit Sub
                'End If

                If SaveUsingParaDetails(pubMeIsNewRecord, objSQL) = 0 Then
                    .Rollback_Trans()
                    Exit Sub
                End If

                Dim i As Integer = 1
                Dim ugR As UltraGridRow
                For Each ugR In UGSOList.Rows

                    If ugR.Hidden = False Then

                        If ugR.Cells("RecID").Value = 0 Then

                            strSQL = "insert into spilRunnSheetDetail (RunnNo,OrderIndex,Status,ThisTimeQty,Comment, geoCoordinations) values " &
                            "(" & txtRunnNumber.Value & "," & ugR.Cells("OrderIndex").Value & "," & DeliveryState.DespatchScheduled & "," &
                            ugR.Cells("ThisTimeDelivery").Value & ",'" & IIf(IsNothing(ugR.Cells("Comment").Value), String.Empty, ugR.Cells("Comment").Value.ToString().Replace("'", "")) &
                            "', '" & If(IsNothing(ugR.Cells("geoCoordinations").Value) = False, ugR.Cells("geoCoordinations").Value, "") & "')"

                            If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                .Rollback_Trans()
                                Exit Sub
                            End If

                            If oProdDef.DeliveryConfirmBy = "RunSheet" Then 'Otherwise this will be done by Delivery Docket

                                strSQL = "update spilInvNum Set Delivery_Status=" & DeliveryState.DespatchScheduled & " " &
                                " where OrderIndex=" & ugR.Cells("OrderIndex").Value & ""

                                If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                    .Rollback_Trans()
                                    Exit Sub
                                End If

                                'strSQL = "update spilInvNumLines Set Delivery_Status=" & DeliveryState.Delivered & " " &
                                '" where OrderIndex=" & ugR.Cells("OrderIndex").Value & ""

                                'If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                '    .Rollback_Trans()
                                '    Exit Sub
                                'End If
                            End If

                            Dim objDocLog As New clsDocumentLogEntry

                            objDocLog.iDocTypeID = ugR.Cells("DocType").Value ' GlassDocTypes.SalesOrder
                            objDocLog.Description1 = "Added to dispatch schedule (" & txtRunnNumber.Value.ToString & ")"
                            objDocLog.iDocID = ugR.Cells("OrderIndex").Value
                            objDocLog.LogAction = "Dispatch Scheduled"
                            objDocLog.DocItemCount = 0
                            objDocLog.DocServiceCount = 0
                            objDocLog.LogDateTime = Now
                            objDocLog.EnteredBy = strUserName
                            objDocLog.Description2 = txtRunnNumber.Value.ToString
                            objDocLog.AddDocLogWithTrans(clsSqlConn.Con, clsSqlConn.Trans)
                            objDocLog = Nothing

                        Else
                            ' update this time qty
                            strSQL = "UPDATE spilRunnSheetDetail Set ThisTimeQty = " & ugR.Cells("ThisTimeDelivery").Value & " , Comment = '" & IIf(IsNothing(ugR.Cells("Comment").Value), String.Empty, ugR.Cells("Comment").Value.ToString().Replace("'", "")) & "' WHERE RecID=" & ugR.Cells("RecID").Value & ""
                            If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                .Rollback_Trans()
                                Exit Sub
                            End If
                        End If

                    Else
                        If ugR.Cells("RecID").Value <> 0 Then

                            strSQL = "delete from spilRunnSheetDetail where RecID=" & ugR.Cells("RecID").Value & ""
                            If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                .Rollback_Trans()
                                Exit Sub
                            End If

                            strSQL = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE RunnNo = " & txtRunnNumber.Value & " AND OrderIndex = " & ugR.Cells("OrderIndex").Value & ""
                            If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                .Rollback_Trans()
                                Exit Sub
                            End If

                            strSQL = "Update spilInvNum set Delivery_Status= " & DeliveryState.UnDelivered & "  where OrderIndex=" & ugR.Cells("OrderIndex").Value & ""
                            If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                .Rollback_Trans()
                                Exit Sub
                            End If

                            Dim objDocLog As New clsDocumentLogEntry
                            objDocLog.iDocTypeID = ugR.Cells("DocType").Value ' GlassDocTypes.SalesOrder
                            objDocLog.Description1 = "Removed from dispatch schedule (" & txtRunnNumber.Value.ToString & ")"
                            objDocLog.iDocID = ugR.Cells("OrderIndex").Value
                            objDocLog.LogAction = "Dispatch Removed"
                            objDocLog.DocItemCount = 0
                            objDocLog.DocServiceCount = 0
                            objDocLog.LogDateTime = Now
                            objDocLog.EnteredBy = strUserName
                            objDocLog.Description2 = txtRunnNumber.Value.ToString
                            objDocLog.AddDocLogWithTrans(.Con, .Trans)
                            objDocLog = Nothing

                        End If
                    End If
                Next

                'Write the record to Appointment table
                ''strSQL = "delete from spil_Scheduler_Appointments where DocumentIndex=" & txtRunnNumber.Value & " and DocumentType=30"
                ''If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                ''    Exit Sub
                ''    .Rollback_Trans()
                ''End If

                ' ''  Dim myDateTime As Date = Format(txtRunningDate.Value, "dd/MM/yyyy") & " " & Format(txtRunnTime.Value, "hh:mm:tt")

                ''strSQL = "set dateformat dmy insert into spil_Scheduler_Appointments (StartDate,EndDate,AllDay,Subject,Location,Description,DocumentIndex,DocumentType,Label) values " & _
                ''            "('" & txtRunnTime.Value & "','" & txtEndTime.Value & "',0,'" & txtRunnNumber.Value.ToString & "'" & _
                ''            ",'" & cboArea.Text & "','" & txtVehRegNo.Text & " (" & txtDrivName.Text & ")" & "'," & txtRunnNumber.Value & ",30,1)"
                ''If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                ''    Exit Sub
                ''    .Rollback_Trans()
                ''End If
                'Write the record to Appointment table

                For Each ugRowLine As UltraGridRow In UGSOLines.Rows()
                    Dim LineDeliveryStatus As Integer

                    If ugRowLine.Cells("OrderQty").Value <= (ugRowLine.Cells("PrevDelQty").Value + ugRowLine.Cells("ThisDelQty").Value + ugRowLine.Cells("RecutQty").Value) Then
                        LineDeliveryStatus = DeliveryState.Delivered
                    ElseIf ugRowLine.Cells("OrderQty").Value > (ugRowLine.Cells("PrevDelQty").Value + ugRowLine.Cells("ThisDelQty").Value + ugRowLine.Cells("RecutQty").Value) Then
                        If (ugRowLine.Cells("PrevDelQty").Value + ugRowLine.Cells("ThisDelQty").Value) > 0 Then
                            LineDeliveryStatus = DeliveryState.PartDelivered
                        Else
                            LineDeliveryStatus = DeliveryState.UnDelivered
                        End If
                    ElseIf (ugRowLine.Cells("PrevDelQty").Value + ugRowLine.Cells("ThisDelQty").Value + ugRowLine.Cells("RecutQty").Value) <= 0 Then
                        LineDeliveryStatus = DeliveryState.UnDelivered
                    End If


                    Dim ScannedBCodes As String = String.Empty

                    'Get scanned barcodes for partial delivery
                    If (objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually) Then
                        Dim scannedBC As New DataSet

                        SQL = "SELECT SerialBarcodeValue FROM spilRunnSheetDownloadedBarCodes WHERE RunnNo = -999 AND Status = 'OK' AND OrderIndex=" & ugRowLine.Cells("OrderIndex").Value & ""
                        scannedBC = .Get_Data_Trans(SQL)

                        For Each Dbc As DataRow In scannedBC.Tables(0).Rows
                            ScannedBCodes = ScannedBCodes & Dbc("SerialBarcodeValue") & "|"
                        Next

                    End If

                    'Update spilRunnSheetDetailLines
                    'If (ugRowLine.Cells("ThisDelQty").Value <> 0) Then
                    If pubMeIsNewRecord = True Then
                        strSQL = "Insert into spilRunnSheetDetailLines " &
                            "(RunnNo,OrderIndex,iInvDetailID,LineNumber,Description,Thickness,Height,Width,GlassWeight,OrderQty," &
                            "PrevDelQty,RecutQty,ThisDelQty,BackOrder,LineTypeID,Barcodes) " &
                            "VALUES (" & txtRunnNumber.Value & "," & ugRowLine.Cells("OrderIndex").Value & "," & ugRowLine.Cells("iInvDetailID").Value & "," &
                            ugRowLine.Cells("LineNo").Value & ",'" & ugRowLine.Cells("Description").Value & "'," & ugRowLine.Cells("Thickness").Value & "," &
                            ugRowLine.Cells("Height").Value & "," & ugRowLine.Cells("Width").Value & "," & ugRowLine.Cells("GlassWeight").Value & "," &
                            ugRowLine.Cells("OrderQty").Value & "," & ugRowLine.Cells("PrevDelQty").Value & "," & ugRowLine.Cells("RecutQty").Value & "," &
                            ugRowLine.Cells("ThisDelQty").Value & "," & ugRowLine.Cells("BackOrder").Value & "," & ugRowLine.Cells("LineTypeID").Value & ",'" & ScannedBCodes & "')"
                    Else
                        strSQL = "UPDATE spilRunnSheetDetailLines SET PrevDelQty= " & ugRowLine.Cells("PrevDelQty").Value & ", " & "RecutQty= " &
                            ugRowLine.Cells("RecutQty").Value & ", ThisDelQty= " & ugRowLine.Cells("ThisDelQty").Value & ", BackOrder=" & IIf(IsDBNull(ugRowLine.Cells("BackOrder").Value), 0, ugRowLine.Cells("BackOrder").Value) &
                            ", Barcodes=') " & ScannedBCodes & "' WHERE RunnNo=" & txtRunnNumber.Value & " AND iInvDetailID=" & ugRowLine.Cells("iInvDetailID").Value & ""
                    End If

                    If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                        .Rollback_Trans()
                        Exit Sub
                    End If


                    'Update Sales Order Line Delivered Qty and Line Status
                    strSQL = "Update spilInvNumLines set fQty_Delivered=fQty_Delivered+" & ugRowLine.Cells("ThisDelQty").Value & ", " &
                        "Delivery_Status=" & LineDeliveryStatus & " where iInvDetailID=" & ugRowLine.Cells("iInvDetailID").Value & ""

                    If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                        .Rollback_Trans()
                        Exit Sub
                    End If

                Next

                'Update RunnNo on despatch scheduled quantity
                If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then

                    strSQL = "UPDATE spilRunnSheetDownloadedBarCodes SET RunnNo = " & txtRunnNumber.Value & " " &
                        "WHERE RunnNo = -999 AND Status = 'AUTO'"
                    If .Exe_Query_Trans(strSQL) = 0 Then
                        .Rollback_Trans()
                        Exit Sub
                    End If

                ElseIf objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then

                    strSQL = "Update spilRunnSheetDownloadedBarCodes set RunnNo = " & txtRunnNumber.Value & " " &
                        "where RunnNo = -999 AND Status = 'OK'"
                    If .Exe_Query_Trans(strSQL) = 0 Then
                        .Rollback_Trans()
                        Exit Sub
                    End If

                End If

                'Update RunnNo on Scanned Trolley Log
                strSQL = "Update spil_DespatchScannedTrolleyBarcodes set RunnNo = " & txtRunnNumber.Value & " " &
                    ", ActivityTypeID = " & TrolleyActivityTypes.TrolleyUse & "  where RunnNo = -999 ;"
                strSQL += "UPDATE FA SET FA.ActivityTypeID = " & TrolleyActivityTypes.TrolleyUse & " FROM spil_DespatchScannedTrolleyBarcodes as ST INNER JOIN spil_fa_Asset as FA ON ST.TrolleyID = FA.AssetId WHERE ST.RunnNo=" & txtRunnNumber.Value & ""
                If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                    .Rollback_Trans()
                    Exit Sub
                End If


                'Get despatched qty from scanned barcodes and sales order header 
                '2016-06-30 for NZ Glass 

                ''bFoundMissingItems = False

                ''Dim iScannedQtyThisTime As Integer = 0
                ''Dim iDeliveredQtyBefore As Integer = 0
                ''Dim iTotalFinishedQtyOnSO As Integer = 0
                ''Dim dsDelQty As DataSet = Nothing
                ''Dim drInv As DataRow = Nothing

                ''For Each ugR In UGSOList.Rows
                ''    If ugR.Hidden = False Then
                ''        strSQL = "select SUM(Qty) as Qty from spilRunnSheetDownloadedBarCodes where RunnNo=" & txtRunnNumber.Value & " and OrderIndex=" & ugR.Cells("OrderIndex").Value
                ''        iScannedQtyThisTime = .Get_ScalerINTEGER_WithTrans(strSQL)


                ''        strSQL = "select TotalFinishedItems, DeliveredFinishedItems from spilInvNum where OrderIndex=" & ugR.Cells("OrderIndex").Value
                ''        dsDelQty = .Get_Data_Trans(strSQL)
                ''        If dsDelQty.Tables(0).Rows.Count > 0 Then
                ''            drInv = dsDelQty.Tables(0).Rows(0)
                ''        End If
                ''        iTotalFinishedQtyOnSO = drInv("TotalFinishedItems")
                ''        iDeliveredQtyBefore = drInv("DeliveredFinishedItems")

                ''        If iScannedQtyThisTime > 0 Then
                ''            strSQL = "Update spilRunnSheetDetail set ThisTimeQty=" & iScannedQtyThisTime & " where RunnNo=" & txtRunnNumber.Value & " and  OrderIndex=" & ugR.Cells("OrderIndex").Value & ""
                ''            If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                ''                Exit Sub
                ''                .Rollback_Trans()
                ''            End If

                ''            If iTotalFinishedQtyOnSO > (iDeliveredQtyBefore + iScannedQtyThisTime) Then
                ''                ugR.Appearance.BackColor = Color.LightPink
                ''                bFoundMissingItems = True
                ''            End If
                ''        Else

                ''        End If

                ''    End If
                ''Next
                '2016-06-30 for NZ Glass 
                'Get despatched qty from scanned barcodes and sales order header

                .Commit_Trans()
                .Con_Close()

                ''If bFoundMissingItems = True Then
                ''    MsgBox("Found missing items on highlighted orders", MsgBoxStyle.Exclamation, "Missing items!")
                ''End If


                tsbPrint.Enabled = True
                tsbUpdate.Enabled = False

                If MsgBox("Successfully updated." & vbCrLf & "Do you want to print the running sheet?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                    strSFormula = "{spilRunnSheetHeader.RunnNo}=" & txtRunnNumber.Value & ""
                    strRepPath = EvoGlassReportPath & "\Reports\Production\Running Sheet.rpt"
                    PrintReport(strSFormula, strRepPath)
                End If


                'If MsgBox("Do you want to print the Delivery Dockets?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                '    frmRunningSheetsList.PrintDeliveryDockets(txtRunnNumber.Value)
                'End If

                SaveMapOnDisk(txtRunnNumber.Value)
                If saveOnly = False Then
                    txtRunnNumber.Value = 0
                    Me.Close()
                End If

            Catch ex As Exception
                WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "SaveRunningSheet")
                MsgBox(ex.Message)
                .Rollback_Trans()
            Finally
                objSQL = Nothing
            End Try
        End With
    End Sub

    Private Sub tsbSelFromList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSelFromList.Click
        Dim objDespatchDef As New clsDespatchDefaults

        If (objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually) Then
            ScanType = RunnSheetScanType.ManualOrder
        ElseIf (objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually) Then
            ScanType = RunnSheetScanType.DesSchPieces
        End If

        frmSOListCommon.frmType = 0     '0-Running sheet     1-Delivery Docket
        frmSOListCommon.WindowState = FormWindowState.Normal
        frmSOListCommon.StartPosition = FormStartPosition.CenterScreen
        frmSOListCommon.ShowDialog()
        GetVehicleDetails(False, False, False)
        'LoadMap()
    End Sub

    Private Sub tsbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSave.Click
        UGSOList.UpdateData()
        UGSOLines.UpdateData()

        '*****Start of running sheet validation part*****
        If (cboArea.Value) = Nothing Or Not IsNumeric(cboArea.Value) Then
            cboArea.Value = Nothing
            MsgBox("Please select correct area.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.DefaultButton1, "Running sheet validation")
            Exit Sub
        End If

        If cboArea.Value = 0 Then
            MsgBox("Please select the area.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.DefaultButton1, "Running sheet validation")
            Exit Sub
        End If

        If cmbFacility.Value = 0 Then
            MsgBox("Please select the branch.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.DefaultButton1, "Running sheet validation")
            Exit Sub
        End If

        If UGSOList.Rows.Count <= 0 Then
            MsgBox("No orders has been added.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.DefaultButton1, "Running sheet validation")
            Exit Sub
        End If

        'If cmbDuration.Value = Nothing Then
        '    MsgBox("Please select a duration.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly + MsgBoxStyle.DefaultButton1, "Running sheet validation")
        '    Exit Sub
        'End If

        Dim IsexceedWeigh As Boolean
        Dim IsexceedHeight As Boolean
        Dim IsexceedWidth As Boolean
        GetVehicleDetails(IsexceedWeigh, IsexceedHeight, IsexceedWidth)

        If IsexceedWeigh AndAlso MsgBox("This order is overweight. Do you want to continue?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.No Then
            Exit Sub
        End If

        If IsexceedHeight AndAlso MsgBox("This order is overheight. Do you want to continue?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.No Then
            Exit Sub
        End If

        If IsexceedWidth AndAlso MsgBox("This order is overwidth. Do you want to continue?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.No Then

            Exit Sub
        End If
        '*****End of running sheet validation part*****

        If MsgBox("Please confirm the update?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
            Call SaveRunningSheet()
        End If
    End Sub

    Private Sub tsbConfirmDeliv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbConfirmDeliv.Click

        If oProdDef.DeliveryConfirmBy = "RunSheet" Then
            If UGSOList.Selected.Rows.Count > 0 Then

                Dim ugR As UltraGridRow
                Dim IsFound As Boolean = False

                For Each ugR In UGSOList.Selected.Rows
                    If ugR.Cells("DelState").Value = DeliveryState.DespatchScheduled Or ugR.Cells("DocState").Value <> GlassProdState.Despatched Then
                        IsFound = True
                        Exit For
                    End If
                Next

                If IsFound = False Then
                    Exit Sub
                End If

                If MsgBox("Please confirm the delivery of selected items?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                    Call UpdateDeliveryStatus()
                End If
            End If
        End If
    End Sub

    Private Sub UpdateDeliveryStatus()
        Dim objDespatchDef As New clsDespatchDefaults
        Dim objSQL As New clsSqlConn
        ''Dim objSQL As New clsInvHeader

        Dim OrderedQty As Double = 0
        Dim PrevDelivereddQty As Double = 0
        Dim ThisTimQty As Double = 0
        Dim DeliveryStatus As Integer = 0

        Dim dsScanQty As DataSet
        Dim drScanQty As DataRow

        With objSQL


            Me.Cursor = Cursors.WaitCursor

            Try
                .Begin_Trans()

                'Check customers assign to trolleys sent out
                strSQL = "SELECT COUNT(RecID) FROM spil_DespatchScannedTrolleyBarcodes WHERE DCLink = 0 AND RunnNo = " & txtRunnNumber.Value & ""
                Dim iNotAssTrolleyCount As Integer = .Get_ScalerINTEGER_WithTrans(strSQL)

                If iNotAssTrolleyCount > 0 Then
                    MessageBox.Show("Please select customer/s to the added trolley/s.")
                    .Rollback_Trans()
                    Exit Sub
                End If

                Dim ugR As UltraGridRow
                Dim IsFound As Boolean = False

                For Each ugR In UGSOList.Selected.Rows

                    If ugR.Cells("DelState").Value = DeliveryState.DespatchScheduled Or ugR.Cells("DocState").Value <> GlassProdState.Despatched Then

                        IsFound = True

                        ''*****Commented By Dasuni - OLD STOCK UPDATE
                        ''If Not (oSODef.DecreaseStockByDelDocket) Then
                        ''    pubMeSODocument = New clsInvHeader(ugR.Cells("OrderIndex").Value, True)
                        ''    objMultiSo.UpdateInventoryStock(oDbCon, pubMeSODocument)
                        ''    pubMeSODocument = Nothing
                        ''End If
                        ''*****Commented By Dasuni - OLD STOCK UPDATE

                        Dim DocLogID As Integer = 0
                        SQL = "SELECT COUNT(DocLogID) FROM spilDocLog WHERE iDocID = " & ugR.Cells("OrderIndex").Value & " AND LogAction = 'Undespatch/Untag'"
                        DocLogID = objSQL.Get_ScalerINTEGER_WithTrans(SQL)

                        If DocLogID = 0 Then

                            ''*****Update Stock and SQM Sold Running Sheet Confirm
                            Dim objInvDefaults As New clsInvDefaults(True)
                            If objInvDefaults.DecreaseStockBy = DecreaseStockBy.Despatch And objInvDefaults.DecreaseStockMethodOnDespatch = DecreaseStockMethodOnDespatch.RunnSheet Then
                                Dim objUpdateStock As New clsUpdateStock
                                Dim InvDetLine As clsInvDetailLine

                                Dim strOrderNum, strClient As String
                                Dim iStockLink As Integer
                                Dim dblThisTimeDelQty As Double
                                Dim DocDate As DateTime

                                Dim InvHeader As New clsInvHeader(ugR.Cells("OrderIndex").Value, True)
                                strOrderNum = txtRunnNumber.Value
                                strClient = InvHeader.cAccountName
                                DocDate = txtRunningDate.Value

                                ''Update Stock
                                For Each InvDetLine In InvHeader.collDetailInvLines
                                    ''Check ThisTimeQuantity <> 0 to update stock
                                    dblThisTimeDelQty = GetThisTimeDelQty(objDespatchDef.RunnSheetScanOption, ugR.Cells("ScanTypeID").Value, InvDetLine, objSQL)
                                    If dblThisTimeDelQty = 0 Then Continue For

                                    ''Check WhseItem = True and AutoReduceStocks = false to update stock
                                    If InvDetLine.SubStockLink <> 0 Then
                                        iStockLink = InvDetLine.SubStockLink
                                    Else
                                        iStockLink = InvDetLine.StockLink
                                    End If
                                    If objUpdateStock.CheckIsWhsAndIsAutoUpdateStock(iStockLink, True) = False Then Continue For

                                    If objUpdateStock.ValidateUpdateStock(InvDetLine, dblThisTimeDelQty, "OUT", strClient, strOrderNum, DocDate) = False Then
                                        MsgBox("Error found while updating stock.", MsgBoxStyle.Information, "Update Error")
                                        .Rollback_Trans()
                                        Exit Sub
                                    End If
                                Next

                                ''Update SQM Sold
                                For Each InvDetLine In InvHeader.collDetailInvLines
                                    ''Check ThisTimeQuantity <> 0 to update stock
                                    dblThisTimeDelQty = GetThisTimeDelQty(objDespatchDef.RunnSheetScanOption, ugR.Cells("ScanTypeID").Value, InvDetLine, objSQL)
                                    If dblThisTimeDelQty = 0 Then Continue For

                                    ''Update SQM Sold only for Glass
                                    If InvDetLine.ItemType = GlassItemTypes.Glass Then
                                        If objUpdateStock.UpdateSQMSold(True, InvDetLine, dblThisTimeDelQty, strClient, strOrderNum, DocDate) = False Then
                                            MsgBox("Error found while updating SQM sold.", MsgBoxStyle.Information, "Update Error")
                                            .Rollback_Trans()
                                            Exit Sub
                                        End If
                                    End If
                                Next

                                ''Update Allocate Qty and SQM
                                If objInvDefaults.IsAllocateStock Then
                                    For Each InvDetLine In InvHeader.collDetailInvLines
                                        ''Check ThisTimeQuantity <> 0 to update stock
                                        dblThisTimeDelQty = GetThisTimeDelQty(objDespatchDef.RunnSheetScanOption, ugR.Cells("ScanTypeID").Value, InvDetLine, objSQL)
                                        If dblThisTimeDelQty = 0 Then Continue For

                                        If objUpdateStock.CalculateReduceAllocateStockAndSQM(InvDetLine, dblThisTimeDelQty) = False Then
                                            MsgBox("Error found while updating reallocate quantity.", MsgBoxStyle.Information, "Update Error")
                                            .Rollback_Trans()
                                            Exit Sub
                                        End If
                                    Next
                                End If
                                ''Update Allocate Qty and SQM

                                objUpdateStock = Nothing
                                InvHeader = Nothing
                            End If
                            objInvDefaults = Nothing
                            ''*****Update Stock and SQM Sold Running Sheet Confirm 
                        End If

                        'Update status as 'Delivered' in spilRunnSheetDetail table
                        strSQL = "update spilRunnSheetDetail set Status = " & DeliveryState.Delivered & " " &
                                " where RecID = " & ugR.Cells("RecID").Value
                        If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                            .Rollback_Trans()
                            Exit Sub
                        End If

                        OrderedQty = 0
                        PrevDelivereddQty = 0
                        ThisTimQty = 0
                        DeliveryStatus = 0

                        'Update ProductionState as 'Despatched' in spilPROD_BATCH table
                        If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
                            OrderedQty = ugR.Cells("TotalGlassPanels").Value
                        Else
                            OrderedQty = ugR.Cells("OrderFinishedItems").Value
                        End If


                        PrevDelivereddQty = ugR.Cells("DeliveredSoFar").Value
                        ThisTimQty = ugR.Cells("ThisTimeDelivery").Value

                        If OrderedQty = ThisTimQty Then
                            DeliveryStatus = DeliveryState.Delivered
                        ElseIf OrderedQty = (PrevDelivereddQty + ThisTimQty) Then
                            DeliveryStatus = DeliveryState.Delivered
                        ElseIf OrderedQty > (PrevDelivereddQty + ThisTimQty) Then
                            DeliveryStatus = DeliveryState.PartDelivered
                        ElseIf PrevDelivereddQty + ThisTimQty = 0 Then
                            DeliveryStatus = DeliveryState.UnDelivered
                        End If

                        If DeliveryStatus = DeliveryState.Delivered Then
                            'Update ProductionState as 'Despatched' in spilInvNum table
                            If ugR.Cells("DocType").Value = GlassDocTypes.SalesOrder Then
                                strSQL = "update spilInvNum set ProductionState = " & GlassProdState.Despatched & " " &
                                " where OrderIndex=" & ugR.Cells("OrderIndex").Value & ""

                                If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                    .Rollback_Trans()
                                    Exit Sub
                                End If
                            ElseIf ugR.Cells("DocType").Value = GlassDocTypes.NCR Then
                                strSQL = "set dateformat dmy update spilInvNum set " &
                                    "ProductionState = " & GlassProdState.Despatched & ", " &
                                    "DocState = " & GlassDocState.Archived & " , InvDate = '" & Now.Date & "', " &
                                    "InvPostedOn = '" & Now & "' where OrderIndex = " & ugR.Cells("OrderIndex").Value & ""

                                If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                    .Rollback_Trans()
                                    Exit Sub
                                End If
                            End If
                        End If

                        'Update fQty_Delivered in spilPROD_BATCH table
                        'Update ProductionState as 'Despatched' in spilInvNumLines table
                        'If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectOrders Then
                        '    strSQL = "UPDATE spilInvNumLines SET fQty_Delivered = fQuantity " &
                        '    "WHERE OrderIndex = " & ugR.Cells("OrderIndex").Value & " AND " &
                        '    "ItemType = " & GlassItemTypes.Glass & ";"
                        '    strSQL += "UPDATE spilPROD_BATCH SET " &
                        '    "spilPROD_BATCH.ProductionState = " & GlassProdState.Despatched & " " &
                        '    "WHERE OrderIndex = " & ugR.Cells("OrderIndex").Value & ";"

                        '    If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                        '        .Rollback_Trans()
                        '        Exit Sub
                        '    End If
                        'Else
                        If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
                            strSQL = "SELECT iInvDetailID, SUM(Qty) AS Qty, BarcodeValue FROM spilRunnSheetDownloadedBarCodes GROUP BY iInvDetailID, OrderIndex, " &
                                    "Status, RunnNo, BarcodeValue HAVING (OrderIndex = " & ugR.Cells("OrderIndex").Value & " AND Status = 'AUTO' " &
                                    "AND RunnNo = " & txtRunnNumber.Value & ")"

                        ElseIf objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
                            If (ugR.Cells("OrderIndex").Value = RunnSheetScanType.ManualOrder) Then
                                strSQL = "SELECT iInvDetailID, fQuantity, BarcodeValue FROM spilInvNumLines WHERE " &
                                    "OrderIndex = " & ugR.Cells("OrderIndex").Value & " AND ((ItemType = " & GlassItemTypes.Glass & " AND M_NO = 0) OR (ItemType = " & GlassItemTypes.Template & " AND M_NO = 0) OR " &
                                    "(ItemType = " & GlassItemTypes.Consumable & " AND M_NO = 0));"

                                'strSQL += "UPDATE spilPROD_BATCH SET " &
                                '"spilPROD_BATCH.ProductionState = " & GlassProdState.Despatched & " " &
                                '"WHERE OrderIndex = " & ugR.Cells("OrderIndex").Value & ";"

                            ElseIf (ugR.Cells("OrderIndex").Value = RunnSheetScanType.ScanPieces) Then
                                strSQL = "SELECT iInvDetailID, SUM(Qty) AS Qty, BarcodeValue FROM spilRunnSheetDownloadedBarCodes GROUP BY iInvDetailID, OrderIndex, " &
                                    "Status, RunnNo, BarcodeValue HAVING (OrderIndex = " & ugR.Cells("OrderIndex").Value & " And Status = 'OK' " &
                                    "AND RunnNo = " & txtRunnNumber.Value & ")"
                            End If
                        End If

                        dsScanQty = .Get_Data_Trans(strSQL)

                        For Each drScanQty In dsScanQty.Tables(0).Rows
                            Dim sFirstLetter As String = Mid(drScanQty("BarcodeValue"), 1, 1)
                            Select Case sFirstLetter
                                Case "L"
                                    strSQL = "UPDATE spilInvNumLines SET fQty_Delivered += " & drScanQty("Qty") & " " &
                                    "WHERE Bar_CodeValue = '" & drScanQty("BarcodeValue") & "';"
                                Case "M"
                                    strSQL = "UPDATE spilInvNumLines SET fQty_Delivered += " & drScanQty("Qty") & " " &
                                    "WHERE MainBar_CodeValue = '" & drScanQty("BarcodeValue") & "';"
                            End Select

                            strSQL += "UPDATE spilPROD_BATCH SET ProductionState = CASE WHEN " &
                                    "Qty_Ord = " & PrevDelivereddQty + ThisTimQty & " THEN " &
                                    "" & GlassProdState.Despatched & " ELSE ProductionState end " &
                                    "where iInvDetailID = " & drScanQty("iInvDetailID") & ";"

                            If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                                .Rollback_Trans()
                                Exit Sub
                            End If
                        Next

                        'End If

                        'Update Delivery_Status, DeliveryDate, DeliveredFinishedItems in spilInvNum table
                        strSQL = "set dateformat dmy update spilInvNum set Delivery_Status = " & DeliveryStatus & ", " &
                              "DeliveryDate='" & Now.Date & "', DeliveredFinishedItems = DeliveredFinishedItems +" & ugR.Cells("ThisTimeDelivery").Value & " where OrderIndex=" & ugR.Cells("OrderIndex").Value & ""
                        If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                            .Rollback_Trans()
                            Exit Sub
                        End If

                        ugR.Cells("DelState").Value = DeliveryStatus
                        ugR.Cells("DocState").Value = GlassProdState.Despatched
                        ugR.Cells("OrdStateName").Value = "Despatched"
                    End If

                Next

                If IsFound = False Then
                    Exit Sub
                End If

                'Update Status as 'Processed' or 'PartProcessed' in spilRunnSheetHeader table
                Dim IsDespatchFound As Boolean = False
                For Each ugR In UGSOList.Rows

                    If (ugR.Cells("DelState").Value = DeliveryState.DespatchScheduled Or ugR.Cells("DocState").Value <> GlassProdState.Despatched) Then
                        IsDespatchFound = True
                        Exit For
                    End If

                Next

                If IsDespatchFound = False Then
                    strSQL = "UPDATE spilRunnSheetHeader set Status=" & GlassReceiptState.Processed & " where RunnNo=" & txtRunnNumber.Value & ""
                Else
                    strSQL = "UPDATE spilRunnSheetHeader set Status=" & GlassReceiptState.PartProcessed & " where RunnNo=" & txtRunnNumber.Value & ""
                End If

                If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                    .Rollback_Trans()
                    Exit Sub
                End If


                'Update Trolley status on Scanned Trolley Log
                strSQL = "UPDATE spil_DespatchScannedTrolleyBarcodes set ActivityTypeID = " & TrolleyActivityTypes.TrolleyOut & " where RunnNo = " & txtRunnNumber.Value & ";"
                strSQL += "UPDATE FA SET FA.ActivityTypeID = " & TrolleyActivityTypes.TrolleyOut & " FROM spil_DespatchScannedTrolleyBarcodes as ST INNER JOIN spil_fa_Asset as FA ON ST.TrolleyID = FA.AssetId WHERE ST.RunnNo=" & txtRunnNumber.Value & ""
                If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                    .Rollback_Trans()
                    Exit Sub
                End If

                'Update Trolley Out Details
                Dim DS_TrolleyOut As DataSet = Nothing
                strSQL = "SELECT TrolleyID,DCLink FROM spil_DespatchScannedTrolleyBarcodes WHERE RunnNo = " & txtRunnNumber.Value & ""
                DS_TrolleyOut = .Get_Data_Trans(strSQL)

                If DS_TrolleyOut.Tables.Count > 0 Then
                    For Each drTrolleyOut As DataRow In DS_TrolleyOut.Tables(0).Rows
                        strSQL = "INSERT INTO spil_fa_TrolleyOutDetails VALUES ( " & drTrolleyOut("TrolleyID") & ", '" & Now.Date & "' , " & drTrolleyOut("DCLink") & " )"
                        If .GET_INSERT_UPDATE("", "", strSQL) = 0 Then
                            .Rollback_Trans()
                            Exit Sub
                        End If
                    Next
                End If

                For Each ugR1 As UltraGridRow In UGSOList.Selected.Rows
                    If AddDeliveryCharge(objSQL, CInt(ugR1.Cells("OrderIndex").Value), CInt(ugR1.Cells("AccountID").Value)) = -1 Then
                        .Rollback_Trans()
                        Exit Sub
                    End If
                Next

                .Commit_Trans()
                .Con_Close()

                'Tagging process
                For Each ugR1 As UltraGridRow In UGSOList.Selected.Rows
                    'If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectOrders Then
                    '    ProcessDespatchTotalOrderQuantity(CInt(ugR1.Cells("OrderIndex").Value))
                    'Else
                    '    ProcessDespatchScanPieces(CInt(ugR1.Cells("OrderIndex").Value))
                    'End If
                    If ugR1.Cells("ScanType").Value = RunnSheetScanType.ManualOrder Then
                        ProcessDespatchTotalOrderQuantity(CInt(ugR1.Cells("OrderIndex").Value))
                    Else
                        ProcessDespatchScanPieces(CInt(ugR1.Cells("OrderIndex").Value))
                    End If
                Next

                MsgBox("Selected item(s) were successfully marked as despatched.", MsgBoxStyle.Information, "Confirmation")

                Dim runnSheetDelCharges As New frmRunnSheetDelCharges
                runnSheetDelCharges.RunnNo = txtRunnNumber.Value
                runnSheetDelCharges.ShowDialog()

                Me.Close()

            Catch ex As Exception
                WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "UpdateDeliveryStatus")
                MsgBox(ex.Message)
                .Rollback_Trans()
                Me.Close()
            Finally
                objSQL = Nothing
                Me.Cursor = Cursors.Default
            End Try
        End With
    End Sub

    Private Function GetThisTimeDelQty(iRunnSheetScanOption As Integer, ScanType As Integer, iInvDetLine As clsInvDetailLine, objSQL As clsSqlConn) As Double
        Dim dblThisTimeDelQty As Double


        If iRunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
            strSQL = "SELECT SUM(Qty) AS Qty FROM spilRunnSheetDownloadedBarCodes WHERE iInvDetailID = " &
                    "" & iInvDetLine.iInvDetailID & " AND Status = 'AUTO' AND RunnNo = " & txtRunnNumber.Value & ""

        ElseIf iRunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then

            strSQL = "SELECT SUM(Qty) AS Qty FROM spilRunnSheetDownloadedBarCodes WHERE iInvDetailID = " &
                    "" & iInvDetLine.iInvDetailID & " AND Status = 'OK' AND RunnNo = " & txtRunnNumber.Value & ""
        End If

        dblThisTimeDelQty = objSQL.Get_ScalerDOUBLE_WithTrans(strSQL)

        Return dblThisTimeDelQty

    End Function

    Private Sub LineAndSerialTagging(BarCodeValue As String, iMachineID As Integer, iDetailID As Integer)
        Try

            Dim myObject As String
testMyText:
            myObject = Mid(BarCodeValue, 1, 1)

            If myObject = vbCr Then
                BarCodeValue = Mid(BarCodeValue, 2)
                GoTo testMyText
            End If

            If myObject = vbLf Then
                BarCodeValue = Mid(BarCodeValue, 2)
                GoTo testMyText
            End If

            didTagged = False

            Dim MainBarcode As String = GetIGUMainBarcode(iDetailID)

            Dim oTag As New clsProductionTagging
            oTag.ScannedBarcode = BarCodeValue
            oTag.TaggedQTY = 1
            oTag.TaggedStationID = iMachineID
            oTag.TaggedMethod = "Despatch"
            oTag.TaggBarcode()

            If Not String.IsNullOrWhiteSpace(MainBarcode) AndAlso MainBarcode.Contains("M") Then
                oTag.TaggingIGUUnit(MainBarcode)
            End If

            oTag = Nothing

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "LineAndSerialTagging")
            MsgBox(ex.Message)
        Finally

        End Try


    End Sub


    Private Function ProcessDespatchTotalOrderQuantity(ByVal InvHeaderID) As Integer
        Dim DS_ITEMS As DataSet = Nothing
        Dim dr1 As DataRow

        Dim objSQL As New clsSqlConn

        Try

            SQL = "SELECT     spilInvNumLines.OrderIndex, spilInvNumLines.idInvoiceLines, spilInvNumLines.fQuantity, spilInvNumLines.iInvDetailID, spilInvNumLines.ItemType " &
                "FROM  spilInvNumLines WITH (NOLOCK) WHERE  spilInvNumLines.OrderIndex = " & InvHeaderID & " ORDER BY spilInvNumLines.idInvoiceLines"
            SQL += " SELECT TOP (1) spilInvNumLines.iInvDetailID, spilInvNumLines.OrderIndex, spilInvNumLines.idInvoiceLines, spilPROD_BATCH.ProcessPath, " &
                "spilPROD_BATCH.STATION_TP_ID FROM spilInvNumLines WITH (NOLOCK) INNER JOIN spilPROD_BATCH WITH (NOLOCK) ON spilInvNumLines.iInvDetailID = spilPROD_BATCH.iInvDetailID " &
                "WHERE(spilInvNumLines.OrderIndex = " & InvHeaderID & ") ORDER BY spilPROD_BATCH.ProcessPath DESC"


            DS_ITEMS = objSQL.GET_INSERT_UPDATE(SQL)

            'this is to get the despatch station for tagging routine
            If DS_ITEMS.Tables(1).Rows.Count > 0 Then
                dr1 = DS_ITEMS.Tables(1).Rows(0)
                wsID = dr1("STATION_TP_ID")
                dr1 = Nothing
            End If


            If DS_ITEMS.Tables(0).Rows.Count > 0 Then


                Dim oProdDef As New clsProdDefaults("")

                If oProdDef.EnablePieceTracking = True Then 'Serial Tagging
                    Dim dsBarcodes As DataSet = Nothing
                    For Each dr1 In DS_ITEMS.Tables(0).Rows
                        If dr1("ItemType") = GlassItemTypes.Glass Then
                            SQL = "select BarCodeV from spilPROD_SERIALS WITH (NOLOCK) where iInvDetailID=" & dr1("iInvDetailID") & " and STATION_TP_ID=" & wsID
                            dsBarcodes = objSQL.GET_INSERT_UPDATE(SQL)
                            For Each drBCode As DataRow In dsBarcodes.Tables(0).Rows
                                Call LineAndSerialTagging(drBCode("BarCodeV").ToString().Trim(), wsID, dr1("iInvDetailID"))
                            Next
                        End If
                    Next

                Else ' >>>>>> Line tagging OLD tagging code
                    For Each dr1 In DS_ITEMS.Tables(0).Rows
                        If dr1("ItemType") = GlassItemTypes.Glass Then
                            ' >>>>>> Line tagging OLD tagging code
                            If Tagging(dr1("OrderIndex"), dr1("iInvDetailID"), dr1("fQuantity"), "Auto") <> 1 Then
                                Return 0
                            End If
                        End If
                    Next
                End If

                '>>> Update despatch flags
                objSQL.Begin_Trans()

                If oProdDef.EnablePieceTracking = False Then 'Line Tagging OLD
                    SQL = "set dateformat dmy Update spilPROD_BATCH set " &
                            "Available=0, Pending=0, Completed=1, Tagged=1, Qty_Out=(Qty_Ord-Qty_Tot_Less), Processed_Date='" & Now & "',Date_Out='" & Now & "' " &
                            "where OrderIndex = " & InvHeaderID & " AND STATION_TP_ID=" & wsID & ""
                    If objSQL.GET_INSERT_UPDATE("", "", SQL) = 0 Then
                        MsgBox("Error in Updating Batch Details", MsgBoxStyle.Critical, "SPIL Glass")
                        objSQL.Rollback_Trans()
                        Return 0
                    End If
                End If

                oProdDef = Nothing

                objSQL.Commit_Trans()
                '>>> Update despatch flags

                Return 1
            Else
                Return 0
            End If

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "ProcessDespatchTotalOrderQuantity")
            MsgBox(ex.Message)
            Return 0
        Finally
            DS_ITEMS.Dispose()
            DS_ITEMS = Nothing
            objSQL = Nothing
        End Try
    End Function

    Private Sub ProcessDespatchScanPieces(OrderIndex As Integer)
        Dim objSQL As New clsSqlConn
        Dim DelStation As Integer = 0
        Dim dsDesSchQty As DataSet
        Dim objDespatchDef As New clsDespatchDefaults

        Try
            strQuery = "SELECT TOP(1) spilPROD_BATCH.STATION_TP_ID FROM spilInvNumLines WITH (NOLOCK) INNER JOIN " &
           "spilPROD_BATCH WITH (NOLOCK) ON spilInvNumLines.iInvDetailID = spilPROD_BATCH.iInvDetailID " &
           "WHERE (spilInvNumLines.OrderIndex = " & OrderIndex & ") ORDER BY spilPROD_BATCH.ProcessPath DESC"

            DelStation = objSQL.Get_ScalerINTEGER(strQuery)
            If DelStation <= 0 Then Exit Sub


            If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
                strSQL = "SELECT COALESCE(NULLIF(SerialBarcodeValue,''), BarcodeValue) AS  BarcodeValue FROM spilRunnSheetDownloadedBarCodes WHERE " &
                    "(OrderIndex = " & OrderIndex & ") AND (Status = 'AUTO') AND (RunnNo = " & txtRunnNumber.Value & ")"

            ElseIf objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
                strSQL = "SELECT COALESCE(NULLIF(SerialBarcodeValue,''), BarcodeValue) AS  BarcodeValue FROM spilRunnSheetDownloadedBarCodes WHERE " &
                    "(OrderIndex = " & OrderIndex & ") AND (Status = 'OK') AND (RunnNo = " & txtRunnNumber.Value & ")"

            End If

            dsDesSchQty = objSQL.GET_DataSet(strSQL)

            For Each drScanQty As DataRow In dsDesSchQty.Tables(0).Rows
                Dim oTag As New clsProductionTagging

                If Mid(drScanQty("BarcodeValue"), 1, 1) = "L" Then
                    oTag.ScannedBarcode = drScanQty("BarcodeValue")
                ElseIf Mid(drScanQty("BarcodeValue"), 1, 1) = "M" Then
                    If ValidateEnteredMainBarcodeIsSerial(drScanQty("BarcodeValue")) Then
                        oTag.ScannedBarcode = Mid(drScanQty("BarcodeValue"), 1, drScanQty("BarcodeValue").LastIndexOf("-"))
                    Else
                        oTag.ScannedBarcode = drScanQty("BarcodeValue")
                    End If
                End If

                oTag.TaggedQTY = 1
                oTag.TaggedStationID = DelStation
                oTag.TaggedMethod = "Despatch"
                oTag.TaggBarcode()
                oTag = Nothing
            Next

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "ProcessDespatchScanPieces")
            MsgBox(ex.Message)
        Finally
            objSQL = Nothing
        End Try
    End Sub

    Private Sub UGSOList_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs)
        Exit Sub

        If pubMeIsNewRecord = True Then

            If MsgBox("Please confirm deletion of selected item(s).", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then

                Dim objSQL As New clsSqlConn

                For Each ugRow As UltraGridRow In e.Rows
                    ugRow.Hidden = True
                    e.Cancel = True

                    strSQL = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE RunnNo = -999 AND OrderIndex = " & ugRow.Cells("OrderIndex").Value & ""
                    If objSQL.Exe_Query(strSQL) = 0 Then
                        Exit Sub
                    End If
                Next

            End If

        Else

            If e.Rows(0).Cells("DelState").Value = DeliveryState.DespatchScheduled Then

                If MsgBox("Please confirm deletion of selected item(s).", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then

                    For Each ugRow As UltraGridRow In e.Rows
                        ugRow.Hidden = True
                        e.Cancel = True
                    Next
                End If

            End If
        End If
    End Sub

    Private Sub ConfirmDeliveryAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfirmDeliveryAllToolStripMenuItem.Click

        If oProdDef.DeliveryConfirmBy = "RunSheet" Then

            If MsgBox("Please confirm", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.No Then Exit Sub

            Dim ugR As UltraGridRow
            Dim IsFound As Boolean = False

            For Each ugR In UGSOList.Rows

                If ugR.Cells("DelState").Value = DeliveryState.DespatchScheduled Or ugR.Cells("DocState").Value <> GlassProdState.Despatched Then
                    ugR.Selected = True
                    IsFound = True
                End If

            Next

            If IsFound = False Then
                Exit Sub
            End If

            Call UpdateDeliveryStatus()

        End If

    End Sub

    Private Sub tsbSelectBCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSelectBCode.Click
        Dim objDespatchDef As New clsDespatchDefaults

        If (objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually) Then
            ScanType = RunnSheetScanType.ManualOrder
        ElseIf (objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually) Then
            ScanType = RunnSheetScanType.DesSchPieces
        End If

        frmRunnSheetTagging.StartPosition = FormStartPosition.CenterScreen
        frmRunnSheetTagging.ShowDialog()
        GetVehicleDetails(False, False, False)
        'LoadMap()
    End Sub

    Private Sub lblAddVehicle_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblAddVehicle.LinkClicked
        frmDeliveryMasterFiles.TC.Tabs(0).Visible = True
        frmDeliveryMasterFiles.TC.Tabs(1).Visible = False
        frmDeliveryMasterFiles.ShowDialog()
        getVehicle()
    End Sub

    Private Sub lblAddDriver_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblAddDriver.LinkClicked
        frmDeliveryMasterFiles.TC.Tabs(0).Visible = False
        frmDeliveryMasterFiles.TC.Tabs(1).Visible = True
        frmDeliveryMasterFiles.ShowDialog()
        getDrivers()
    End Sub

    Private Sub cmbDuration_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cmbDuration.ValueChanged
        txtEndTime.Value = DateAdd(DateInterval.Minute, cmbDuration.Value, txtRunnTime.Value)
    End Sub

    Private Sub txtRunnTime_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtRunnTime.ValueChanged
        txtRunningDate.Value = CDate(txtRunnTime.Value)
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        oProdDef = New clsProdDefaults("")
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub tsbSelectByPieceBCode_Click(sender As System.Object, e As System.EventArgs) Handles tsbSelectByPieceBCode.Click
        ScanType = RunnSheetScanType.ScanPieces

        frmRunnSheetScanPieces.StartPosition = FormStartPosition.CenterScreen
        If pubMeIsNewRecord = True Then
            frmRunnSheetScanPieces.iRunnSheetNo = -1
        Else
            frmRunnSheetScanPieces.iRunnSheetNo = txtRunnNumber.Value
        End If
        frmRunnSheetScanPieces.ShowDialog()
        GetVehicleDetails(False, False, False)
        'LoadMap()
    End Sub

    Private Sub UGSOList_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        'Dim odrQty = If(IsDBNull(e.Cell.Row.Cells("OrderFinishedItems").Value), 0, e.Cell.Row.Cells("OrderFinishedItems").Value)
        'Dim thiQty = If(IsDBNull(e.Cell.Row.Cells("ThisTimeDelivery").Value), 0, e.Cell.Row.Cells("ThisTimeDelivery").Value)
        'If (thiQty < thiQty) Then
        '    e.Cell.Row.Cells("ThisTimeDelivery").Appearance.BackColor = Color.Yellow
        '    e.Cell.Row.Cells("ThisTimeDelivery").Appearance.ForeColor = Color.Red
        'End If
    End Sub

    Private Sub tsbViewProdStatus_Click(sender As Object, e As EventArgs) Handles tsbViewProdStatus.Click
        If UGSOList.Selected.Rows.Count > 0 Then
            If IsPieceTrackingEnabled = True Then
                frmProd_FlowPieces.iOrderNumber = UGSOList.Selected.Rows(0).Cells("OrderIndex").Value
                frmProd_FlowPieces.lblOrdNo.Text = "Order Status - " & UGSOList.Selected.Rows(0).Cells("OrderNum").Value.ToString & " (" & CDate(UGSOList.Selected.Rows(0).Cells("OrderDate").Value.ToString) & ")"
                frmProd_FlowPieces.StartPosition = FormStartPosition.CenterScreen
                frmProd_FlowPieces.ShowDialog()
            Else
                frmProd_Flow.iOrderNumber = UGSOList.Selected.Rows(0).Cells("OrderIndex").Value
                frmProd_Flow.lblOrdNo.Text = "Order Status - " & UGSOList.Selected.Rows(0).Cells("OrderNum").Value.ToString & " (" & CDate(UGSOList.Selected.Rows(0).Cells("OrderDate").Value.ToString) & ")"
                frmProd_Flow.StartPosition = FormStartPosition.CenterScreen
                frmProd_Flow.ShowDialog()
            End If
        End If
    End Sub

    Private Sub tsbVeiwOrder_Click(sender As Object, e As EventArgs) Handles tsbVeiwOrder.Click
        If UGSOList.Selected.Rows.Count > 0 Then

            If UGSOList.Selected.Rows(0).Cells("DocType").Value = GlassDocTypes.ARCreditNote Or UGSOList.Selected.Rows(0).Cells("DocType").Value = GlassDocTypes.ARDebitNote Or
                UGSOList.Selected.Rows(0).Cells("DocType").Value = GlassDocTypes.Receipt Or
                UGSOList.Selected.Rows(0).Cells("DocType").Value = GlassDocTypes.POSInvoice Or UGSOList.Selected.Rows(0).Cells("DocType").Value = GlassDocTypes.SOInvoiced Then
                Exit Sub
            End If

            Dim mySalesDoc As New frmSO

            mySalesDoc.pubMeIsNewRecord = False
            mySalesDoc.pubMeCalledBy = GlassDocCalledBy.EditDocListDoubleClicked

            mySalesDoc.pubMeOrderIndex = UGSOList.Selected.Rows(0).Cells("OrderIndex").Value
            mySalesDoc.pubMeSpilDocTypeID = UGSOList.Selected.Rows(0).Cells("DocType").Value

            mySalesDoc.WindowState = FormWindowState.Minimized

            mySalesDoc.LoadDocumentMasterData()
            mySalesDoc.OpenSO()
            mySalesDoc.SetDocumentControlProperties()

            mySalesDoc.StartPosition = FormStartPosition.CenterScreen
            mySalesDoc.Show()

            mySalesDoc.Refresh()

            mySalesDoc.WindowState = FormWindowState.Maximized
            mySalesDoc.BringToFront()

        End If
    End Sub

    Private Sub lblAddArea_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblAddArea.LinkClicked
        Dim frmAreas As New frmCustomerAreas()
        frmAreas.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        frmAreas.Height = 400
        frmAreas.Width = 600
        frmAreas.ShowIcon = True
        frmAreas.ShowDialog()
        GET_AREAS()
    End Sub

    Private Sub tsbScanTrolley_Click(sender As Object, e As EventArgs) Handles tsbScanTrolley.Click
        Dim objScanTrolley As New frmRunnSheetScanTrolley
        objScanTrolley.iRunnMode = pubMeMode
        objScanTrolley.iRunnNo = txtRunnNumber.Value
        objScanTrolley.ShowDialog()
    End Sub

    Private Sub tsbScanOptions_Click(sender As Object, e As EventArgs) Handles tsbScanOptions.Click
        Dim objScanOptions As New frmRunnSheetScanOptions
        objScanOptions.ShowDialog()

        VisibleSelectByPieceBCode()
    End Sub

    Private Sub tsmRefreshQuantities_Click(sender As Object, e As EventArgs) Handles tsmRefreshQuantities.Click
        Dim objSQL As New clsSqlConn
        Dim objDespatchDef As New clsDespatchDefaults

        objSQL.Begin_Trans()

        strQuery = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE RunnNo = -999 AND Status = 'AUTO'"
        If objSQL.Exe_Query_Trans(strQuery) = 0 Then
            objSQL.Rollback_Trans()
            Exit Sub
        End If

        objSQL.Commit_Trans()

        For Each ugR In UGSOList.Rows
            Dim iThisTimeQty As Integer
            If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
                iThisTimeQty = GetDespathScheduledQuantity(ugR.Cells("OrderIndex").Value)
            End If

            If pubMeMode = 1 Then 'New
                ugR.Cells("ThisTimeDelivery").Value = iThisTimeQty
            ElseIf pubMeMode = 2 Then 'Edit
                ugR.Cells("ThisTimeDelivery").Value += iThisTimeQty
            End If
        Next
    End Sub

    Private Function ValidateEnteredMainBarcodeIsSerial(strBarcode As String) As Boolean
        Dim sFirstLetter As String = Mid(strBarcode, 1, 1)
        Dim rgx As New System.Text.RegularExpressions.Regex("^M-\d+-\d+-\d+$") '' Match Main barcode 
        Dim match As System.Text.RegularExpressions.Match = rgx.Match(strBarcode)
        If sFirstLetter = "M" And match.Success Then
            Return True
        Else
            Return False
        End If
    End Function





    Private Sub DELETEToolStripMenuItem_Click(sender As Object, e As EventArgs)

        If txtRunnNumber.Value <> 0 Then

            If MsgBox("Are you sure you want to delete", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.No Then
                Exit Sub
            End If
        Else
            If Me.UGSOList.Selected.Rows.Count > 0 Then

                Me.UGSOList.DeleteSelectedRows()
            Else

                MessageBox.Show("There are no rows selected. Select rows first.")
            End If

        End If
    End Sub



    Private Sub AddCommentToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DeleteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem1.Click
        If Me.UGSOList.Selected.Rows.Count > 0 Then
            'Me.UGSOList.DeleteSelectedRows()
            DeleteRowsAdded()
        End If

    End Sub

    Private Sub AddCommentToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles AddCommentToolStripMenuItem.Click
        If Me.UGSOList.Selected.Rows.Count > 0 Then
            Dim frmComment As New frmComment()
            frmComment.Comment = UGSOList.Selected.Rows(0).Cells("Comment").Value
            frmComment.ShowDialog()
            If frmComment.DialogResult = Windows.Forms.DialogResult.OK Then
                For Each ugR As UltraGridRow In UGSOList.Selected.Rows
                    ugR.Cells("Comment").Value = frmComment.Comment
                Next
            End If
        End If
    End Sub


    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        DeleteToolStripMenuItem1_Click(sender, e)
    End Sub

    Private Function GetIGUMainBarcode(iDetailID As Integer) As String
        Dim o As New clsSqlConn
        Dim sMainBarcode As String = o.Get_ScalerString("select MainBar_CodeValue from spilInvNumLines WITH (NOLOCK) where iInvDetailID=" & iDetailID & "")
        o = Nothing
        Return sMainBarcode
    End Function

    Private Function AddDeliveryCharge(ByVal objSQL As clsSqlConn, OrderIndex As Integer, Customer As Integer) As Integer
        Dim DelChargeTypeId As Integer = 0
        Dim DelChargeRate As Double = 0
        Dim LineID As Integer = 0
        Dim InvHeaderID As Integer = 0
        Try

            Dim InvHeader As New clsInvHeader(OrderIndex, True)

            SQL = "SELECT DeliveryChargeID,FLRate FROM spilFuelLevRates WHERE FLID IN(SELECT FLID FROM client WHERE DCLink=" & InvHeader.AccountID & ")"
            Dim DSDel As New DataSet
            DSDel = InvHeader.Get_Data_Trans(SQL)

            For Each dr As DataRow In DSDel.Tables(0).Rows
                DelChargeTypeId = CInt(dr("DeliveryChargeID"))
                DelChargeRate = CDbl(dr("FLRate"))
            Next

            If DelChargeTypeId = DeliveryChargeType.RunningSheet And InvHeader.DocType <> GlassDocTypes.NCR Then

                If Customer = lastAccountId Then
                    LogDeliveryCharge(InvHeader, DelChargeRate, objSQL, 0)
                    Exit Function
                End If

                LogDeliveryCharge(InvHeader, DelChargeRate, objSQL, 1)

                InvHeader.InvTotExcl += DelChargeRate
                InvHeader.InvTotTax = (InvHeader.InvTotExcl * InvHeader.TaxRate) / 100
                InvHeader.InvTotIncl = InvHeader.InvTotExcl + InvHeader.InvTotTax
                InvHeader.OrderTotal = InvHeader.InvTotIncl
                InvHeaderID = InvHeader.AddHeader()

                If InvHeaderID = -1 Then
                    Return -1
                End If

                LineID = AddDeliveryLine(oSoDefaults.RunningSheetDelChargeStockLink, InvHeader, DelChargeRate, objSQL)

                If LineID = -1 Then
                    Return -1
                End If

                lastAccountId = Customer

            End If

            Return 1
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "AddDeliveryCharge")
            Return -1
        End Try
    End Function

    Public Function AddDeliveryLine(ByVal StockLink As Integer, ByVal InvHeader As clsInvHeader, ByVal DelChargeRate As Double, ByVal objSQL As clsSqlConn) As Integer
        Try
            'Dim InvLines As New clsInvDetailLine
            Dim objDocLog As New clsDocumentLogEntry
            Dim LineDocID As Integer = 0
            Dim client As New clsCustomer()
            client.DCLink = InvHeader.AccountID
            client.GetCustomerDataFromOpenConnection()

            objDocLog.iDocTypeID = InvHeader.DocType
            objDocLog.LogAction = "Added"
            objDocLog.iDocID = InvHeader.OrderIndex
            objDocLog.LogDateTime = Now
            objDocLog.EnteredBy = strUserName ' "EDI"

            'SQL = "SELECT     StockLink, Code, Description_1, uiIIPRICETYPEID, uiIISRVPRICEID FROM StkItem WHERE StockLink = " & StockLink & ""

            SQL = "SELECT StkItem.Description_1, StkItem.cSimpleCode as Code, StkItem.StockLink,  StkItem.ufIIThickness, TaxRate.TaxRate,stkItem.uiIISRVPRICEID, " &
          " StkItem.TTI, StkItem.uiIIPRICETYPEID, StkItem.uiIIItemType,StkItem.Description_3,StkItem.AddDetails FROM StkItem LEFT OUTER JOIN " &
          " TaxRate ON StkItem.TTI = TaxRate.Code " &
          " WHERE StockLink = " & StockLink & ""

            With objSQL
                Dim DelLine As New DataSet
                DelLine = .Get_Data_Trans(SQL)
                SQL = "SELECT (MAX(idInvoiceLines)+1) AS idInvoiceLines FROM  spilInvNumLines WHERE OrderIndex=" & InvHeader.OrderIndex
                Dim idInvoiceLines As Integer = .Get_ScalerINTEGER_WithTrans(SQL)

                For Each mDr As DataRow In DelLine.Tables(0).Rows

                    Dim clsInvDetLine As clsInvDetailLine = New clsInvDetailLine
                    clsInvDetLine.iInvDetailID = 0
                    clsInvDetLine.LN = 0 'drLines("LN")
                    clsInvDetLine.M_NO = 0
                    clsInvDetLine.OrderIndex = InvHeader.OrderIndex
                    clsInvDetLine.UniqueLN = idInvoiceLines
                    clsInvDetLine.idInvoiceLines = idInvoiceLines
                    clsInvDetLine.ProductionState = GlassProdState.None
                    clsInvDetLine.LineType = LineState.Normal
                    clsInvDetLine.ProcessedID = GlassInvLineProductionState.UnProcessed '<<< to identify the row has been cut/processsed
                    clsInvDetLine.StockLink = CInt(mDr("StockLink"))
                    clsInvDetLine.cSimpleCode = mDr("Code")
                    clsInvDetLine.cDescription = mDr("Description_1")
                    clsInvDetLine.Description_1 = mDr("Description_1")
                    clsInvDetLine.fQuantity = 1
                    clsInvDetLine.iHeight = 0
                    clsInvDetLine.iWidth = 0
                    clsInvDetLine.fVolume = 0
                    clsInvDetLine.fThickness = 0
                    clsInvDetLine.bToughened = 0
                    clsInvDetLine.Measure = mDr("uiIISRVPRICEID")
                    clsInvDetLine.Method = 0
                    clsInvDetLine.Unit = 1
                    clsInvDetLine.MainItem = 1
                    clsInvDetLine.ItemType = GlassItemTypes.Service
                    clsInvDetLine.ItemTypeCategory = 0
                    clsInvDetLine.iTaxTypeID = 0
                    clsInvDetLine.fTaxRate = mDr("TaxRate")
                    clsInvDetLine.TaxCode = mDr("TTI")
                    clsInvDetLine.IsPriceItem = 0 ' True 'IIf((DR.Cells("IsPriceItem").Value) = True, 1, 0)
                    clsInvDetLine.PRICE_TYPES_ID = mDr("uiIIPRICETYPEID")
                    clsInvDetLine.fUnitCost = DelChargeRate
                    clsInvDetLine.PriceList = 0
                    clsInvDetLine.PriceCategory = "T"
                    clsInvDetLine.OrgPrice = DelChargeRate
                    clsInvDetLine.fItem_SC = 0
                    clsInvDetLine.fOriginal_Price = clsInvDetLine.fUnitCost
                    If clsInvDetLine.fUnitCost <= 0 Or clsInvDetLine.Unit <= 0 Then
                        clsInvDetLine.fItem_Net = 0
                        clsInvDetLine.fItem_tax = 0
                        clsInvDetLine.fItem_Gross = 0
                    Else
                        clsInvDetLine.fItem_Net = Math.Round(((clsInvDetLine.fUnitCost - clsInvDetLine.fDiscount_Amount) * clsInvDetLine.Unit), 2, MidpointRounding.AwayFromZero)
                        clsInvDetLine.fItem_tax = Math.Round((clsInvDetLine.fItem_Net * (clsInvDetLine.fTaxRate / 100)), 2, MidpointRounding.AwayFromZero)
                        clsInvDetLine.fItem_Gross = Math.Round((clsInvDetLine.fItem_Net + clsInvDetLine.fItem_tax), 2, MidpointRounding.AwayFromZero)
                    End If

                    clsInvDetLine.fTotal_Amt = 0
                    objDocLog.DocItemNetAmt = objDocLog.DocItemNetAmt + clsInvDetLine.fItem_Net
                    objDocLog.DocItemGSTAmt = objDocLog.DocItemGSTAmt + clsInvDetLine.fItem_tax
                    clsInvDetLine.fService_Net = 0
                    clsInvDetLine.fService_tax = 0
                    clsInvDetLine.fService_Gross = 0
                    clsInvDetLine.fDiscount_Amount = 0
                    clsInvDetLine.fDiscount_Percen = 0
                    clsInvDetLine.fSTD_COST = 0
                    clsInvDetLine.Bar_CodeValue = ""
                    clsInvDetLine.MainBar_CodeValue = ""
                    clsInvDetLine.Delivery_Status = 0
                    clsInvDetLine.fQty_Delivered = 0
                    clsInvDetLine.fQty_Invoiced = 0
                    clsInvDetLine.LineNotes = ""
                    clsInvDetLine.Comment = ""
                    clsInvDetLine.Comment2 = ""
                    clsInvDetLine.LineComments = ""
                    clsInvDetLine.Motif = 0
                    clsInvDetLine.fCompletedQuantity = 0
                    clsInvDetLine.ReBatchOriginID = 0
                    clsInvDetLine.ReBatchReasonID = 0
                    clsInvDetLine.ReBatchQty = 0
                    clsInvDetLine.ReBatchRefLine = 0
                    clsInvDetLine.ReBatchStationID = 0
                    clsInvDetLine.TemplItemPercentage = 0
                    clsInvDetLine.Qty_Suspended = 0
                    clsInvDetLine.FacilityIDCurrent = 1
                    clsInvDetLine.FacilityIDNext = 0
                    clsInvDetLine.Current_Staion_TP_ID = 0
                    clsInvDetLine.Next_Staion_TP_ID = 0
                    clsInvDetLine.FacilityID = 1
                    clsInvDetLine.SubStockLink = 0
                    clsInvDetLine.ReservedQty = 0
                    clsInvDetLine.QtyOnSO = 0
                    clsInvDetLine.IsExternalItem = 0
                    clsInvDetLine.H1 = 0
                    clsInvDetLine.ShapeFileName = ""
                    clsInvDetLine.TempInvLineDetID = 0

                    clsInvDetLine.Foreign_InvTotExcl = Math.Round(clsInvDetLine.fItem_Net * client.ExRate, 2, MidpointRounding.AwayFromZero)
                    clsInvDetLine.Foreign_InvTotTax = Math.Round(clsInvDetLine.fItem_tax * client.ExRate, 2, MidpointRounding.AwayFromZero)
                    clsInvDetLine.Foreign_InvTotIncl = Math.Round(clsInvDetLine.fItem_Gross * client.ExRate, 2, MidpointRounding.AwayFromZero)

                    InvHeader.AddInvDetailLines(clsInvDetLine)
                    clsInvDetLine = Nothing
                    objDocLog.DocItemCount = 1

                    LineDocID = InvHeader.UpdateInvDetailLines()
                    If LineDocID = -1 Then
                        InvHeader.Rollback_Trans()
                        Exit Function
                    End If

                    objDocLog.Quoted = 0
                    objDocLog.QuotedNetAmt = 0
                    objDocLog.QuotedGSTAmt = 0
                    '' objDocLog.DocNetAmt = frmNullInvoiceDetails.ExclAmt
                    '' objDocLog.DocGSTAmt = frmNullInvoiceDetails.GSTAmt
                    objDocLog.DocDelChrgNetAmt = DelChargeRate
                    objDocLog.DocDelChrgGSTAmt = 0
                    objDocLog.Description1 = "Delivery Charge Added"
                    objDocLog.Description2 = ""

                    If objDocLog.AddDocLogWithTrans(clsInvHeader.Con, clsInvHeader.Trans) = 0 Then
                        InvHeader.Rollback_Trans()
                        Exit Function
                    End If

                Next

            End With
            Return 1
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "AddDeliveryLine")
            Return -1
        End Try
    End Function

    Public Function LogDeliveryCharge(ByVal InvHeader As clsInvHeader, ByVal DelChargeRate As Double, ByVal objSQL As clsSqlConn, IsChargeAdded As Integer) As Integer
        Try
            SQL = "SELECT RunnNo FROM  spilRunnSheetDetail WHERE OrderIndex=" & InvHeader.OrderIndex
            Dim RunnNo As Integer = objSQL.Get_ScalerINTEGER_WithTrans(SQL)

            SQL = "INSERT INTO spilRunnSheetDelCharge(OrderIndex,RunnNo,DelRate,IsChargeAdded)VALUES(" & InvHeader.OrderIndex & "," & RunnNo & "," & DelChargeRate & "," & IsChargeAdded & ")"

            If objSQL.Exe_Query_Trans(SQL) = 0 Then
                Return -1
            End If

            Return 1
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "LogDeliveryCharge")
            Return -1
        End Try
    End Function

    'Added by Hashini on 26-04-2019 - Comment section of runsheets (Citiwest)
    Private Sub tsb_ClearComment_Click(sender As Object, e As EventArgs) Handles tsb_ClearComment.Click
        Dim ugR As UltraGridRow = UGSOList.Selected.Rows(0)
        ugR.Cells("Comment").Value = String.Empty
    End Sub

    'Added by Hashini on 26-04-2019 - Unable to remove orders from run (Citiwest)
    Private Sub DeleteRowsAdded()
        Try
            If UGSOList.Selected.Rows.Count = 0 Then
                Exit Sub
            End If

            Dim objSQL As New clsSqlConn

            If pubMeIsNewRecord = True Then
                If MsgBox("Please confirm deletion of selected item(s).", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                    'Dim objSQL As New clsSqlConn

                    For Each ugRow As UltraGridRow In UGSOList.Selected.Rows
                        ugRow.Hidden = True

                        strSQL = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE RunnNo = -999 AND OrderIndex = " & ugRow.Cells("OrderIndex").Value & ""
                        If objSQL.Exe_Query(strSQL) = 0 Then
                            Exit Sub
                        End If

                        For Each ugLine As UltraGridRow In UGSOLines.Rows
                            If (ugRow.Cells("OrderIndex").Value = ugLine.Cells("OrderIndex").Value) Then
                                ugLine.Delete(False)
                            End If
                        Next
                    Next
                    btnRefreshGoogleMap.Visible = True
                End If
            Else
                'If UGSOList.Selected.Rows(0).Cells("DelState").Value = DeliveryState.DespatchScheduled Then
                '    If MsgBox("Please confirm deletion of selected item(s).", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then

                '        For Each ugRow As UltraGridRow In UGSOList.Selected.Rows
                '            ugRow.Hidden = True
                '        Next
                '    End If
                'End If

                For Each ugRow As UltraGridRow In UGSOList.Selected.Rows
                    If ugRow.Cells("DelState").Value = DeliveryState.DespatchScheduled Or ugRow.Cells("DelState").Value = DeliveryState.UnDelivered Then
                        If MsgBox("Selected order(s) will be permanently removed from this Running sheet. Please confirm.", MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton1, "Confirmation to delete") = MsgBoxResult.Yes Then

                            strSQL = "delete from spilRunnSheetDetail where RecID=" & ugRow.Cells("RecID").Value & ";"
                            strSQL += "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE RunnNo = " & txtRunnNumber.Value & " AND OrderIndex = " & ugRow.Cells("OrderIndex").Value & ";"
                            strSQL += "Update spilInvNum set Delivery_Status= " & DeliveryState.UnDelivered & "  where OrderIndex=" & ugRow.Cells("OrderIndex").Value & ";"

                            If objSQL.Exe_Query(strSQL) = 0 Then
                                MsgBox("Invalid deletion. Please try again.", "Confirmation")
                                Exit Sub
                            End If

                            For Each ugLine As UltraGridRow In UGSOLines.Rows
                                If (ugRow.Cells("OrderIndex").Value = ugLine.Cells("OrderIndex").Value) Then
                                    strSQL = "delete from spilRunnSheetDetailLines where iInvDetailID=" & ugLine.Cells("iInvDetailID").Value & ""

                                    If objSQL.Exe_Query(strSQL) = 0 Then
                                        MsgBox("Invalid deletion. Please try again.", "Confirmation")
                                        Exit Sub
                                    End If

                                    ugLine.Delete(False)
                                End If
                            Next

                            Dim objDocLog As New clsDocumentLogEntry
                            objDocLog.iDocTypeID = ugRow.Cells("DocType").Value ' GlassDocTypes.SalesOrder
                            objDocLog.Description1 = "Removed from dispatch schedule (" & txtRunnNumber.Value.ToString & ")"
                            objDocLog.iDocID = ugRow.Cells("OrderIndex").Value
                            objDocLog.LogAction = "Dispatch Removed"
                            objDocLog.DocItemCount = 0
                            objDocLog.DocServiceCount = 0
                            objDocLog.LogDateTime = Now
                            objDocLog.EnteredBy = strUserName
                            objDocLog.Description2 = txtRunnNumber.Value.ToString
                            objDocLog.AddLogRecord()
                            objDocLog = Nothing

                            'UGSOList.Selected.Rows(0).Delete(False)
                            ugRow.Delete(False)
                        End If
                    End If
                Next
                btnRefreshGoogleMap.Visible = True
            End If
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "DeleteRowsAdded")
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Location Extensions"
#Region "Location Variables"
    Dim directionURL As String = ""
    Dim stricMapURL As String = ""
    Dim staticMapimage As Image
    Dim GoogleMapAPIExtensionsObj As New GoogleMapAPIExtensions
#End Region
#Region "Google Map CRUD"
    'Function SaveLocationDetails(ByRef clsInvHeaderObj As clsInvHeader) As Integer
    '    Dim collspPara As New Collection
    '    Dim colPara As New spParameters
    '    Dim newSQLQuery As String = ""

    '    Try
    '        colPara.ParaName = "@GoogleMapAPIExtensionsObj"
    '        colPara.ParaValue = If(IsNothing(geometryLocation) = False, geometryLocation, "")
    '        collspPara.Add(colPara)

    '        colPara.ParaName = "@OrderIndex"
    '        colPara.ParaValue = clsInvHeaderObj.OrderIndex
    '        collspPara.Add(colPara)

    '        newSQLQuery += " UPDATE [dbo].[spilRunnSheetHeader] SET [GoogleMapAPIExtensionsObj] = @GoogleMapAPIExtensionsObj WHERE OrderIndex= @OrderIndex"

    '        Return clsInvHeaderObj.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara)
    '    Catch ex As Exception
    '        modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
    '        Return 0
    '    End Try

    'End Function
#End Region
#Region "Location Funtions"
    Function GetLocationDataList() As List(Of String)
        Dim waypoints As List(Of String) = New List(Of String)
        Try
            For Each rows As UltraGridRow In UGSOList.Rows
                waypoints.Add((rows.Cells("geoCoordinations").Value))
            Next
            Return waypoints
        Catch ex As Exception
            Return waypoints
        End Try
    End Function

    Function GetMapURL() As Integer
        Dim waypoints As List(Of String) = New List(Of String)
        Dim urlString As String
        Dim urlArray() As String
        Try
            If isGoogleAPIActive = False Then
                Exit Function
            End If
            waypoints = GetLocationDataList()
            If UGSOList.Rows.Count > 0 AndAlso waypoints.Count > 0 Then
                urlString = GoogleMapAPIExtensionsObj.GetGoogleMapURL(waypoints, googleAPIKey, If(showLocationInMinimap = False, GoogleMapURLType.DirectionURL, GoogleMapURLType.All))
                urlArray = urlString.Split("|")
                directionURL = urlArray(0)
                stricMapURL = urlArray(1)
                If IsNothing(directionURL) = True Then
                    directionURL = ""
                End If
                If IsNothing(stricMapURL) = True Then
                    stricMapURL = ""
                End If
                If stricMapURL <> "" Then
                    pbGoogleMap.Load(stricMapURL)
                    staticMapimage = pbGoogleMap.Image
                End If
                Return 1
            Else
                modGlazingQuoteExtension.GQShowMessage("No address to show the map", Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
                Return 0
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
            Return 0
        End Try
    End Function

    Sub ShowDeliveryRouteMapURL()
        Try
            'Dim frmGlazingAddessLocatorGmapObj As New frmGlazingAddessLocatorGmap()
            'frmGlazingAddessLocatorGmapObj.WebBrowser1.Navigate(directionURL)
            'frmGlazingAddessLocatorGmapObj.WebBrowser1.Dock = DockStyle.Fill
            'frmGlazingAddessLocatorGmapObj.WebBrowser1.Visible = True
            'frmGlazingAddessLocatorGmapObj.ShowDialog()

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Sub EmailDeliveryRouteMap()
        Dim emailDefaultSubject As String = ""
        Dim emailDefaultBody As String
        Dim completeEmailBody As String
        Dim fullPath As String
        Dim fullPathArray() As String
        Dim AttDocPara As eMailAttachDocumentPara = Nothing
        Dim collAttDocuments As New Collection
        Try
            If GoogleMapURLValidator(directionURL) = False Then
                Exit Sub
            End If
            emailDefaultBody = GoogleMapAPIExtensionsObj.GetEmailDetails(emailDefaultSubject)
            fullPath = GetImageLocationPath(txtRunnNumber.Value)
            fullPathArray = fullPath.Split("|")
            completeEmailBody = GoogleMapAPIExtensionsObj.SetCompleteEmailBodyString(directionURL, emailDefaultBody, "src='cid:" & fullPathArray(1) & "'", SetEmailContetendgridDetails())
            AttDocPara.DocumentName = fullPathArray(1)
            AttDocPara.DocumentPath = fullPathArray(0) & fullPathArray(1)
            collAttDocuments.Add(AttDocPara)
            If GoogleMapAPIExtensionsObj.SendAnEmail(emailDefaultSubject, completeEmailBody, txtDrivName.Value, collAttDocuments) = 1 Then

            Else
                modGlazingQuoteExtension.GQShowMessage("Email not sent", Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)

            End If

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Sub ShowDeliveryMap()
        Try
            If isGoogleAPIActive = False Then
                Exit Sub
            End If
            Dim defaultBroswer As String = GoogleMapAPIExtensionsObj.GetSystemDefaultBrowser()
            If IsNothing(defaultBroswer) = False AndAlso defaultBroswer <> "" Then
                If GetMapURL() = 1 Then
                    Process.Start(defaultBroswer, directionURL)
                    ShowDeliveryRouteMapURL()
                End If
            Else
                modGlazingQuoteExtension.GQShowMessage("No web browser found.", Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Sub EmilDeliveryMap()
        Try
            If UGSOList.Rows.Count > 0 Then
                If GoogleMapURLValidator(directionURL) = False Then
                    Exit Sub
                End If
                If IsNothing(txtDrivName.Value) = False AndAlso txtDrivName.Value <> "" Then
                    Dim result As DialogResult = modGlazingQuoteExtension.GQShowMessage("Do you wont to save the delivery sheet", Me.Text, MsgBoxStyle.Question, "question", "Save before email.")
                    If result = DialogResult.Yes Then
                        SaveRunningSheet(True)
                        EmailDeliveryRouteMap()
                    End If
                Else
                    modGlazingQuoteExtension.GQShowMessage("Please check the selected driver email address.", Me.Text, MsgBoxStyle.Critical, "warning", "Error in the driver's email address.")
                End If
            Else
                modGlazingQuoteExtension.GQShowMessage("No address to show the map", Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Function GetImageLocationPath(ByRef runningSheetNumber As Integer, Optional ByRef getCombination As Boolean = True) As String
        Dim imagePath As String
        Dim imageFileName As String
        Dim fulPath As String
        Try
            imagePath = EvoGlassReportPath & "\Google_Static_Maps"
            imageFileName = "imgGSM" & runningSheetNumber & ".png"
            If Directory.Exists(imagePath) = False Then
                Directory.CreateDirectory(imagePath)
            End If
            imagePath = imagePath & "\"
            If getCombination = True Then
                fulPath = imagePath & "|" & imageFileName
            Else
                fulPath = imagePath & imageFileName
            End If
            Return fulPath
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
            Return ""
        End Try
    End Function

    Function SaveMapOnDisk(ByRef runningSheetNumber As Integer) As Integer
        Dim imageFullPath As String
        Dim mapImage As Image
        Try
            mapImage = pbGoogleMap.Image
            If IsNothing(mapImage) = True Then
                GetMapURL()
            End If
            imageFullPath = GetImageLocationPath(runningSheetNumber, False)

            If IsNothing(imageFullPath) = False Then
                GoogleMapAPIExtensionsObj.SaveRouteMap(imageFullPath, mapImage)
            Else

            End If
            Return 1
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
            Return 0
        End Try
    End Function

    Function SetEmailContetendgridDetails() As String
        Dim list As String = ""
        Dim sqlQuary As String = ""
        Dim ds As DataSet
        Dim orderIndexList As New List(Of Integer)
        Try
            For Each row As UltraGridRow In UGSOList.Rows
                orderIndexList.Add(row.Cells("OrderIndex").Value)
            Next

            For Each item As Integer In GoogleMapAPIExtensionsObj.waypointOrder
                sqlQuary += "Select  OrderIndex, OrderNum, Address1, Address2, Address3, Address4, Address5 from spilInvNum where OrderIndex = '" & orderIndexList(item) & "'"
            Next
            ds = GoogleMapAPIExtensionsObj.GetData(sqlQuary)
            For Each tb As DataTable In ds.Tables
                For Each dr As DataRow In tb.Rows
                    Dim wayPoints As String = If(dr("Address1") <> "", dr("Address1") & ", ", "") & If(dr("Address2") <> "", dr("Address2") & ", ", "") & If(dr("Address3") <> "", dr("Address3") & ", ", "") & If(dr("Address4") <> "", dr("Address4") & ", ", "") & dr("Address5") & " (" & dr("orderNum") & ")"
                    list += "<li style='font-size: 10px; line-height: 10px; text-align: Left;'><span style='font-size: 10px; line-height: 20px;'>" & wayPoints & "</span></li>"
                Next
            Next
            Return list
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Function GoogleMapURLValidator(ByRef GoogleMapURL As String) As Boolean
        Try
            If IsNothing(GoogleMapURL) = True Then
                GoogleMapURL = ""
            ElseIf GoogleMapURL = "" Then
                If GetMapURL() = 0 Then
                    Return True
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region
#Region "Location Controllers handlers"
    Sub GoogleMapControllerHandler(ByRef state As Boolean)
        Try
            TsbShowDeliveryRouteMapToolStripMenuItem.Enabled = state
            TsbEmailDeliveryRouteMapToolStripMenuItem.Enabled = state
            btnRefreshGoogleMap.Visible = state
            pbGoogleMap.Visible = state
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
        End Try
    End Sub

    Private Sub TsbShowDeliveryRouteMapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TsbShowDeliveryRouteMapToolStripMenuItem.Click
        ShowDeliveryMap()
    End Sub

    Private Sub TsbEmailDeliveryRouteMapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TsbEmailDeliveryRouteMapToolStripMenuItem.Click
        EmilDeliveryMap()
    End Sub

    Private Sub TsbLocationSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TsbLocationSettingsToolStripMenuItem.Click
        Try
            Dim frmEmailSettings As New frmGoogleMapAPIDashboard()
            frmEmailSettings.StartPosition = FormStartPosition.CenterScreen
            frmEmailSettings.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub
#End Region
#End Region

    'Get SO line when click on SO
    Private Sub UGSOList_ClickCell(sender As Object, e As ClickCellEventArgs) Handles UGSOList.ClickCell
        Dim ugSORow As UltraGridRow = UGSOList.ActiveRow
        lblOrderNum.Text = ugSORow.Cells("OrderNum").Value.ToString()
        For Each ugRow As UltraGridRow In UGSOLines.Rows
            If ugRow.Cells("OrderIndex").Value = ugSORow.Cells("OrderIndex").Value Then
                ugRow.Hidden = False
            Else
                ugRow.Hidden = True
            End If
        Next
    End Sub

    'Show production for the selected SO
    Private Sub tsb_ShowProductionStatus_Click(sender As Object, e As EventArgs) Handles tsb_ShowProductionStatus.Click
        frmProd_FlowPieces.iOrderNumber = UGSOList.Selected.Rows(0).Cells("OrderIndex").Value
        frmProd_FlowPieces.lblOrdNo.Text = "Order Status - " & UGSOList.Selected.Rows(0).Cells("OrderNum").Value.ToString & " (" & CDate(UGSOList.Selected.Rows(0).Cells("OrderDate").Value.ToString) & ")"
        frmProd_FlowPieces.StartPosition = FormStartPosition.CenterScreen
        frmProd_FlowPieces.ShowDialog()
    End Sub

    Private Sub GetVehicleDetails(ByRef IsexceedWeigh As Boolean, ByRef IsexceedHeight As Boolean, ByRef IsexceedWidth As Boolean)
        Dim objSQL As New clsSqlConn
        Dim drRow As DataRow = Nothing
        Try
            lblTruckWeight.Text = "0KG"
            lblHeight.Text = "0"
            lblWidth.Text = "0"
            If IsNothing(txtVehRegNo.Value) Then Exit Sub

            If IsNumeric(txtVehRegNo.Value) Then
                SQL = "SELECT * FROM  dbo.spilVehicleMaster WHERE ID=" & txtVehRegNo.Value
            Else
                SQL = "SELECT * FROM  dbo.spilVehicleMaster WHERE Name='" & txtVehRegNo.Value & "'"
            End If


            Dim dsDetsils As DataSet = objSQL.GET_DATA_SQL(SQL)
            If dsDetsils.Tables.Count > 0 Then
                If dsDetsils.Tables(0).Rows.Count > 0 Then
                    drRow = dsDetsils.Tables(0).Rows(0)
                End If
            End If
            Dim MaxGlassWeight As Decimal = 0
            Dim MaxHeight As Decimal = 0
            Dim MaxWidth As Decimal = 0

            If Not IsNothing(drRow) Then
                MaxGlassWeight = IIf(IsDBNull(drRow("MaxGlassWeight")), 0, drRow("MaxGlassWeight"))
                MaxHeight = IIf(IsDBNull(drRow("MaxHeight")), 0, drRow("MaxHeight"))
                MaxWidth = IIf(IsDBNull(drRow("MaxWidth")), 0, drRow("MaxWidth"))
                If MaxGlassWeight > 0 Then
                    MaxGlassWeight = MaxGlassWeight * 1000  ''1 t = 1000 kg
                    lblTruckWeight.Text = MaxGlassWeight.ToString() & "KG"
                End If

                If MaxHeight > 0 Then
                    lblHeight.Text = MaxHeight
                End If

                If MaxWidth > 0 Then
                    lblWidth.Text = MaxWidth
                End If

            End If

            Dim GlassWeight As Decimal = 0
            For Each druRow As UltraGridRow In UGSOLines.Rows.GetRowEnumerator(GridRowType.DataRow, Nothing, Nothing)
                GlassWeight += IIf(IsDBNull(druRow.Cells("GlassWeight").Value), 0, druRow.Cells("GlassWeight").Value)
                Dim Height As Decimal = IIf(IsDBNull(druRow.Cells("Height").Value), 0, druRow.Cells("Height").Value)
                Dim Width As Decimal = IIf(IsDBNull(druRow.Cells("Width").Value), 0, druRow.Cells("Width").Value)

                If Height > 0 AndAlso MaxHeight > 0 Then
                    If Height > MaxHeight Then
                        IsexceedHeight = True
                    End If
                End If

                If Width > 0 AndAlso MaxWidth > 0 Then
                    If Height > MaxWidth Then
                        IsexceedWidth = True
                    End If
                End If
            Next

            If MaxGlassWeight > 0 AndAlso GlassWeight > 0 Then
                If GlassWeight > MaxGlassWeight Then
                    IsexceedWeigh = True
                End If
            End If

            If IsexceedWeigh Then
                lblTruckWeight.ForeColor = Color.Red
            Else
                lblTruckWeight.ForeColor = Color.Black
            End If

            lblTotalWeight.Text = Math.Round(GlassWeight, 2)

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "GetVehicleDetails")
            MsgBox(ex.Message)
        Finally
            objSQL = Nothing
        End Try
    End Sub

    Private Function GetGlassWeight(ByRef ThisTimeDelQty As Integer, ByVal drRow As DataRow) As Decimal
        Try
            If IsNothing(drRow) Then Return 0

            Dim GlassWeight As Decimal = 0

            If drRow("ItemType") = GlassItemTypes.Glass Then
                Dim Height As Decimal = IIf(IsDBNull(drRow("Height")), 0, drRow("Height"))
                Dim Width As Decimal = IIf(IsDBNull(drRow("Width")), 0, drRow("Width"))
                Dim Thickness As Decimal = IIf(IsDBNull(drRow("Thickness")), 0, drRow("Thickness"))
                If Height > 0 AndAlso Width > 0 AndAlso Thickness > 0 Then
                    GlassWeight = ((((Height * Width) / 1000000) * ThisTimeDelQty) * Thickness) * 2.5
                End If
            End If

            Return GlassWeight

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "GetGlassWeight")
            Throw ex
        End Try
        Return 0
    End Function

    Private Sub txtVehRegNo_ValueChanged(sender As Object, e As EventArgs) Handles txtVehRegNo.ValueChanged
        GetVehicleDetails(False, False, False)
    End Sub

    Public Sub AddOrderData(ByVal iOrderIndex As Integer, Optional iInvDetailID As Integer = 0)
        Dim objDespatchDef As New clsDespatchDefaults
        Dim objSQL As New clsSqlConn
        Dim DS_ITEMS As New DataSet
        Dim dr1 As DataRow
        Dim ugR As UltraGridRow = Nothing
        Dim collspPara As New Collection
        Dim colPara As New spParameters
        'Dim strWhere As String
        Dim iThisTimeQty As Integer = 0
        Dim iDeliveredAndPendingQty As Integer = 0
        Dim iOrderQuantity As Integer = 0
        Dim booIsFound As Boolean = False

        Try
            SQL = "SELECT RunnNo, (SELECT OrderNum FROM spilInvNum WHERE OrderIndex = " & iOrderIndex & ") AS OrderNum FROM spilRunnSheetDetail WHERE OrderIndex = " & iOrderIndex & ""
            Dim dsPrevRuunNos As DataSet = objSQL.GET_DataSet(SQL)

            If dsPrevRuunNos.Tables(0).Rows.Count > 0 Then
                If MsgBox("This order (" & dsPrevRuunNos.Tables(0).Rows(0)("OrderNum") & ") has already been added to running sheet no: " & dsPrevRuunNos.Tables(0).Rows(0)("RunnNo") & ". Do you want to continue?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "SPIL Glass") = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If

            'If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
            '    iThisTimeQty = GetDespathScheduledQuantity(iOrderIndex)
            '    If iThisTimeQty <= 0 Then Exit Sub
            'ElseIf objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
            '    iThisTimeQty = GetManuallyScanPieces(iOrderIndex)
            '    If iThisTimeQty <= 0 Then Exit Sub
            'End If

            If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
                iThisTimeQty = GetDespathScheduledQuantity(iOrderIndex)
                If iThisTimeQty <= 0 Then Exit Sub
            ElseIf objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
                If (ScanType = RunnSheetScanType.ScanPieces) Then
                    iThisTimeQty = GetManuallyScanPieces(iOrderIndex)
                    If iThisTimeQty <= 0 Then Exit Sub
                End If
            End If

            colPara.ParaName = "@OrderIndex"
            colPara.ParaValue = iOrderIndex
            collspPara.Add(colPara)

            objSQL.GET_DataSetFromSP("sp_Spil_RunningSheetSOandSoLines", collspPara, DS_ITEMS)

            For Each ugR In UGSOList.Rows
                If ugR.Cells("OrderIndex").Value = iOrderIndex Then
                    booIsFound = True
                    If iInvDetailID = 0 Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next

            'If iInvDetailID = 0 Then
            '    For Each ugR In UGSOList.Rows
            '        If ugR.Cells("OrderIndex").Value = iOrderIndex Then
            '            booIsFound = True
            '            Exit Sub
            '        End If
            '    Next
            'Else
            '    For Each ugR In UGSOLines.Rows
            '        If ugR.Cells("OrderIndex").Value = iOrderIndex AndAlso ugR.Cells("iInvDetailID").Value = iInvDetailID Then
            '            booIsFound = True
            '            Exit For
            '        End If
            '    Next

            '    If booIsFound Then
            '        For Each ugR In UGSOList.Rows
            '            If ugR.Cells("OrderIndex").Value = iOrderIndex Then
            '                booIsFound = True
            '                Exit For
            '            End If
            '        Next
            '    End If

            'End If

            For Each dr1 In DS_ITEMS.Tables(0).Rows
                If booIsFound = False Then
                    ugR = UGSOList.DisplayLayout.Bands(0).AddNew
                End If

                ugR.Cells("RecID").Value = 0
                ugR.Cells("OrderIndex").Value = dr1("OrderIndex")
                ugR.Cells("DocState").Value = 0     'dr1("InvDocState")
                ugR.Cells("RunnStatus").Value = 0
                ugR.Cells("DelState").Value = 0
                ugR.Cells("OrdStateName").Value = dr1("ProdState")
                ugR.Cells("OrderNum").Value = dr1("OrderNum")
                ugR.Cells("AccountID").Value = dr1("AccountID")
                ugR.Cells("CustomerName").Value = dr1("Customer")
                ugR.Cells("OrderDate").Value = dr1("OrderDate")
                ugR.Cells("ExtOrderNum").Value = dr1("CustOrdNo")
                ugR.Cells("iAreasID").Value = 0     'dr1("CustOrdNo")
                ugR.Cells("AreaName").Value = dr1("Area")
                ugR.Cells("DueDate").Value = dr1("DueDate")
                ugR.Cells("DocType").Value = dr1("DocType")
                ugR.Cells("OrderFinishedItems").Value = dr1("OrderFinishedItems")
                ugR.Cells("TotalGlassPanels").Value = dr1("TotalGlassPanels")
                ugR.Cells("geoCoordinations").Value = dr1("geoCoordinations")
                ugR.Cells("ScanTypeID").Value = ScanType

                If (ScanType = RunnSheetScanType.ManualOrder) Then
                    ugR.Cells("ScanType").Value = "Manual"
                ElseIf (ScanType = RunnSheetScanType.ScanPieces) Then
                    ugR.Cells("ScanType").Value = "Scan"
                ElseIf (ScanType = RunnSheetScanType.DesSchPieces) Then
                    ugR.Cells("ScanType").Value = "Despatch Pieces"
                End If

                If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.AutoSelectDesSchPieces Then
                    iOrderQuantity = ugR.Cells("TotalGlassPanels").Value
                Else
                    iOrderQuantity = ugR.Cells("OrderFinishedItems").Value
                End If

                iDeliveredAndPendingQty = FillDeliveredSoFarQty(iOrderIndex, ugR.Cells("RunnStatus").Value)
                If iOrderQuantity >= iDeliveredAndPendingQty Then
                    ugR.Cells("DeliveredSoFar").Value = iDeliveredAndPendingQty
                Else
                    ugR.Cells("DeliveredSoFar").Value = iOrderQuantity
                End If

                'If RunnSheetScanningOption = RunnSheetScanOptions.AutoSelectOrders Then
                '    ugR.Cells("ThisTimeDelivery").Value = ugR.Cells("OrderFinishedItems").Value - ugR.Cells("DeliveredSoFar").Value
                'Else
                '    If iOrderQuantity >= CInt(ugR.Cells("DeliveredSoFar").Value) + iThisTimeQty Then
                '        ugR.Cells("ThisTimeDelivery").Value = iThisTimeQty
                '    Else
                '        ugR.Cells("ThisTimeDelivery").Value = iOrderQuantity - CInt(ugR.Cells("DeliveredSoFar").Value)
                '    End If
                'End If

                If ugR.Cells("ScanTypeID").Value = RunnSheetScanType.ManualOrder Then
                    ugR.Cells("ThisTimeDelivery").Value = ugR.Cells("OrderFinishedItems").Value - ugR.Cells("DeliveredSoFar").Value
                Else
                    If iOrderQuantity >= CInt(ugR.Cells("DeliveredSoFar").Value) + iThisTimeQty Then
                        ugR.Cells("ThisTimeDelivery").Value = iThisTimeQty
                    Else
                        ugR.Cells("ThisTimeDelivery").Value = iOrderQuantity - CInt(ugR.Cells("DeliveredSoFar").Value)
                    End If
                End If

                If (CInt(ugR.Cells("ThisTimeDelivery").Value) + CInt(ugR.Cells("DeliveredSoFar").Value)) < iOrderQuantity Then
                    ugR.Appearance.ForeColor = Color.FromArgb(130, 7, 7)
                Else
                    ugR.Appearance.ForeColor = Color.FromArgb(4, 73, 6)
                End If

                ugR.Cells("BalanceToDeliver").Value = CInt(ugR.Cells("TotalGlassPanels").Value) - (CInt(ugR.Cells("DeliveredSoFar").Value) + CInt(ugR.Cells("ThisTimeDelivery").Value))

                If (ugR.Cells("DocType").Value = GlassDocTypes.NCR) Then
                    ugR.Appearance.ForeColor = Color.SaddleBrown
                End If

                Call AddAllSalesOrderLines(iOrderIndex, DS_ITEMS.Tables(1), iInvDetailID, objDespatchDef, ugR.Cells("ScanTypeID").Value)
            Next

            'Dim ugSORow As UltraGridRow = UGSOList.ActiveRow
            'For Each ugRow As UltraGridRow In UGSOLines.Rows
            '    If ugRow.Cells("OrderIndex").Value = ugSORow.Cells("OrderIndex").Value Then
            '        ugRow.Hidden = False
            '    Else
            '        ugRow.Hidden = True
            '    End If
            'Next
            UGSOList.Rows(0).Selected = False
            UGSOLines.ActiveRow = Nothing
            UGSOList.ActiveRow.Appearance.BackColor = Color.Transparent

            Dim ugSORow As UltraGridRow = UGSOList.ActiveRow
            For Each ugRow As UltraGridRow In UGSOLines.Rows
                If ugRow.Cells("OrderIndex").Value = ugSORow.Cells("OrderIndex").Value Then
                    ugRow.Hidden = False
                Else
                    ugRow.Hidden = True
                End If
            Next

            lblOrderNum.Text = ugSORow.Cells("OrderNum").Value.ToString()

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "AddOrderData")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        Finally
            dr1 = Nothing
            DS_ITEMS = Nothing
            objSQL = Nothing
        End Try
    End Sub

    'Added by Hashini on 18-10-2018 - To display SO Lines
    Private Sub AddAllSalesOrderLines(ByVal iOrderIndex As Integer, ByRef dtSOAllLines As DataTable, ByVal iInvDetailID As Integer, objDespatchDef As clsDespatchDefaults, ScanType As Integer)
        Dim objSQL As New clsSqlConn
        Try

            Dim boolsFound As Boolean = False
            Dim ugRow1 As UltraGridRow = Nothing

            For Each ugRow1 In UGSOLines.Rows
                If ugRow1.Cells("OrderIndex").Value = iOrderIndex AndAlso ugRow1.Cells("iInvDetailID").Value = iInvDetailID Then
                    boolsFound = True
                    Exit For
                End If
            Next

            'For Each drSOLine As DataRow In DS_ITEMS.Tables(0).Rows
            For Each drSOLine As DataRow In dtSOAllLines.Rows
                If boolsFound = False Then
                    ugRow1 = UGSOLines.DisplayLayout.Bands(0).AddNew()
                End If

                ugRow1.Cells("OrderIndex").Value = drSOLine("OrderIndex")
                ugRow1.Cells("iInvDetailID").Value = drSOLine("DetailID")
                ugRow1.Cells("LineNo").Value = drSOLine("idInvoiceLines")
                ugRow1.Cells("Description").Value = drSOLine("ItemDescription")
                ugRow1.Cells("Thickness").Value = drSOLine("Thickness")
                ugRow1.Cells("Height").Value = drSOLine("Height")
                ugRow1.Cells("Width").Value = drSOLine("Width")
                ugRow1.Cells("Size").Value = drSOLine("Height") & " X " & drSOLine("Width")

                ugRow1.Cells("OrderQty").Value = drSOLine("fQuantity")
                ugRow1.Cells("PrevDelQty").Value = drSOLine("DeliveredFinishedItems")
                ugRow1.Cells("RecutQty").Value = drSOLine("ReBatchQty")

                ugRow1.Cells("ThisDelQty").Activation = Activation.AllowEdit

                If objDespatchDef.RunnSheetScanOption = RunnSheetScanOptions.ScanPiecesManually Then
                    If (ScanType = RunnSheetScanType.ManualOrder) Then
                        If drSOLine("fQuantity") - drSOLine("DeliveredFinishedItems") > 0 Then
                            ugRow1.Cells("ThisDelQty").Value = drSOLine("fQuantity") - drSOLine("DeliveredFinishedItems") - drSOLine("ReBatchQty")
                        Else
                            ugRow1.Cells("ThisDelQty").Value = 0
                        End If
                    ElseIf (ScanType = RunnSheetScanType.ScanPieces) Then
                        ugRow1.Cells("ThisDelQty").Value = GetManuallyScanLinePieces(iOrderIndex, drSOLine("DetailID"))
                    End If
                Else
                    If drSOLine("fQuantity") - drSOLine("DeliveredFinishedItems") > 0 Then
                        ugRow1.Cells("ThisDelQty").Value = drSOLine("fQuantity") - drSOLine("DeliveredFinishedItems") - drSOLine("ReBatchQty")
                    Else
                        ugRow1.Cells("ThisDelQty").Value = 0
                    End If
                End If

                ugRow1.Cells("GlassWeight").Value = GetGlassWeight(ugRow1.Cells("ThisDelQty").Value, drSOLine)

                ugRow1.Cells("LineType").Value = drSOLine("LineType")
                ugRow1.Cells("LineTypeID").Value = drSOLine("LineTypeID")
                ugRow1.Cells("Barcodes").Hidden = True


                'ugRow.Cells("BackOrder").Value = 0

                'If drSOLine("LineType") = 1 Then
                '    ugRow1.Cells("LineType").Value = "Re-Cut"
                '    ugRow1.CellAppearance.ForeColor = Color.Red
                'ElseIf drSOLine("LineType") = 2 Then
                '    ugRow1.Cells("LineType").Value = "NCR"
                '    ugRow1.CellAppearance.ForeColor = Color.Magenta
                'Else
                '    ugRow1.Cells("LineType").Value = "Normal"
                'End If
                'GetBarcodes(ugRow1, iOrderIndex, drSOLine("DetailID"), boolsFound)

                If (ugRow1.Cells("ThisDelQty").Value + ugRow1.Cells("PrevDelQty").Value + ugRow1.Cells("RecutQty").Value) < ugRow1.Cells("OrderQty").Value Then
                    ugRow1.CellAppearance.ForeColor = Color.FromArgb(130, 7, 7)
                    ugRow1.Cells("ThisDelQty").Appearance.BackColor = Color.Yellow
                Else
                    ugRow1.CellAppearance.ForeColor = Color.Green
                End If

                If drSOLine("LineTypeID") = LineState.ReBatched Then
                    ugRow1.CellAppearance.ForeColor = Color.Red

                    'ElseIf drSOLine("LineTypeID") = LineState.NCR Then
                    '    ugRow1.CellAppearance.ForeColor = Color.Magenta
                End If

            Next
            'SOlinesBalanceToDeliver()
            FillBackOrdQty()

            ugRow1 = Nothing
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "AddAllSalesOrderLines")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        End Try
    End Sub

    'Update this time delivery qty for scan barcode
    Public Sub UpdateThisTimeDelivery(ByRef dtSOLine As DataTable)
        Try
            'ugSOLines.DisplayLayout.Bands(0).Columns("ThisDelQty").CellActivation = Activation.NoEdit

            If (UGSOLines.Rows.Count > 0) Then
                Dim ugR As UltraGridRow
                For Each ugR In UGSOLines.Rows
                    ugR.Cells("ThisDelQty").Value = 0
                    ' If ugR.Cells("iInvDetailID").Value = iLineDetailID Then
                    For Each drSOLine As DataRow In dtSOLine.Rows
                        If ugR.Cells("iInvDetailID").Value = drSOLine("iInvDetailID") Then
                            'If ugR.Cells("ThisDelQty").Value > 0 Then
                            'If drSOLine("fQuantity") - (drSOLine("fQty_Delivered") + ugR.Cells("ThisDelQty").Value) > 0 Then
                            ugR.Cells("ThisDelQty").Value = ugR.Cells("ThisDelQty").Value + 1
                            '    Else
                            '        ugR.Cells("ThisDelQty").Value = 0
                            '    End If
                            'Else
                            '    If drSOLine("fQuantity") - drSOLine("fQty_Delivered") > 0 Then
                            '        ugR.Cells("ThisDelQty").Value = 1
                            '    Else
                            '        ugR.Cells("ThisDelQty").Value = 0
                            '    End If
                            'End If
                        End If
                    Next
                    ' End If

                    'If (ugR.Cells("ItemType").Value = GlassItemTypes.Consumable Or ugR.Cells("ItemType").Value = GlassItemTypes.Service) Then
                    ugR.Cells("ThisDelQty").Activation = Activation.AllowEdit

                    If CInt(ugR.Cells("ThisDelQty").Value) < ugR.Cells("OrderQty").Value Then
                        ugR.Cells("ThisDelQty").Appearance.BackColor = Color.Yellow
                        ugR.Appearance.ForeColor = Color.Red
                    End If

                Next
            End If

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "UpdateThisTimeDelivery")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        End Try
    End Sub

    ''Update Balance to deliver for SO lines
    'Private Sub SOlinesBalanceToDeliver()
    '    Try
    '        If (UGSOLines.Rows.Count > 0) Then
    '            Dim ugR As UltraGridRow
    '            For Each ugR In UGSOLines.Rows
    '                If ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value) < 0 Then
    '                    ugR.Cells("BalanceToDeliver").Value = 0
    '                Else
    '                    ugR.Cells("BalanceToDeliver").Value = ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value + ugR.Cells("RecutQty").Value)
    '                End If
    '            Next
    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Update back qty for scan barcode
    Private Sub FillBackOrdQty()
        Try
            If (UGSOLines.Rows.Count > 0) Then
                Dim ugR As UltraGridRow
                For Each ugR In UGSOLines.Rows
                    If ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value) < 0 Then
                        ugR.Cells("BackOrder").Value = 0
                    Else
                        ugR.Cells("BackOrder").Value = ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value + +ugR.Cells("RecutQty").Value)
                    End If
                Next
            End If

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "FillBackOrdQty")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        End Try
    End Sub

    Private Function GetManuallyScanLinePieces(OrderIndex As Integer, InvDetailID As Integer) As Integer
        Dim objSQL As New clsSqlConn
        Dim strQuery As String
        Dim sDesSchBarcodes As New ArrayList
        Dim iScanPieces As Integer = 0
        Dim dsScanPieces As DataSet

        objSQL.Begin_Trans()

        strQuery = "DELETE FROM spilRunnSheetDownloadedBarCodes WHERE SerialBarcodeValue " &
            "IN (SELECT BarCodeV FROM spilPROD_SERIALS WHERE Qty_ReBatched = 1 AND " &
            "OrderIndex = " & OrderIndex & " AND iInvDetailID = " & InvDetailID & ") AND Status = 'OK'"
        If objSQL.Exe_Query_Trans(strQuery) = 0 Then
            objSQL.Rollback_Trans()
            Exit Function
        End If

        strQuery = "SELECT COUNT(BarcodeValue) FROM spilRunnSheetDownloadedBarCodes " &
           "WHERE OrderIndex = " & OrderIndex & " AND iInvDetailID = " & InvDetailID & " AND Status = 'OK' AND RunnNo = -999"

        iScanPieces = objSQL.Get_ScalerINTEGER_WithTrans(strQuery)

        objSQL.Commit_Trans()

        Return iScanPieces
    End Function

    'Private Sub GetBarcodes(ByVal ugRow As UltraGridRow, ByVal iOrderIndex As Integer, ByVal iInvDetailId As Integer, boolsFound As Boolean)
    '    Try
    '        Dim objSQL As New clsSqlConn
    '        Dim URow As UltraGridRow

    '        If IsPieceTrackingEnabled = True Then
    '            SQL = "SELECT BarCodeV AS BarCode FROM spilPROD_SERIALS WHERE OrderIndex = " & iOrderIndex & " AND iInvDetailID = " & iInvDetailId & " AND ProcessPath = 1"
    '        Else
    '            SQL = "SELECT Bar_CodeValue AS BarCode FROM spilInvNumLines WHERE OrderIndex = " & iOrderIndex & " AND iInvDetailID = " & iInvDetailId & ""
    '        End If

    '        Dim dsBarcodes As DataSet = objSQL.GET_DATA_SQL(SQL)

    '        cmb_Barcodes.DataSource = dsBarcodes
    '        cmb_Barcodes.ValueMember = "BarCode"
    '        cmb_Barcodes.DisplayMember = "BarCode"

    '        cmb_Barcodes.CheckedListSettings.ItemCheckArea = ItemCheckArea.Item
    '        cmb_Barcodes.CheckedListSettings.EditorValueSource = EditorWithComboValueSource.CheckedItems
    '        cmb_Barcodes.CheckedListSettings.ListSeparator = " / "

    '        If Not cmb_Barcodes.DisplayLayout.Bands(0).Columns.Exists("Selected") Then
    '            Dim c As UltraGridColumn = Me.cmb_Barcodes.DisplayLayout.Bands(0).Columns.Add()

    '            c.Key = "Selected"
    '            c.Header.Caption = String.Empty
    '            c.Header.CheckBoxVisibility = HeaderCheckBoxVisibility.Always
    '            c.DataType = GetType(Boolean)
    '            c.Header.VisiblePosition = 0

    '            Me.cmb_Barcodes.CheckedListSettings.CheckStateMember = "Selected"
    '            Me.cmb_Barcodes.CheckedListSettings.EditorValueSource = Infragistics.Win.EditorWithComboValueSource.CheckedItems
    '            Me.cmb_Barcodes.CheckedListSettings.ListSeparator = " / "
    '            Me.cmb_Barcodes.CheckedListSettings.ItemCheckArea = Infragistics.Win.ItemCheckArea.Item

    '        End If

    '        Dim dsScannedBrcodes As DataSet

    '        If RunnSheetScanningOption = RunnSheetScanOptions.AutoSelectDesSchPieces Or RunnSheetScanningOption = RunnSheetScanOptions.ScanPiecesManually Then
    '            SQL = "SELECT SerialBarcodeValue FROM spilRunnSheetDownloadedBarCodes WHERE OrderIndex = " & iOrderIndex & " AND iInvDetailID = " & iInvDetailId & " AND RunnNo = -999"
    '            dsScannedBrcodes = objSQL.GET_DATA_SQL(SQL)
    '        End If


    '        For Each ur As UltraGridRow In cmb_Barcodes.Rows
    '            If RunnSheetScanningOption = RunnSheetScanOptions.DespatchTotalOrderQuantity Then
    '                ur.Cells("Selected").Value = True
    '            ElseIf RunnSheetScanningOption = RunnSheetScanOptions.AutoSelectDesSchPieces Or RunnSheetScanningOption = RunnSheetScanOptions.ScanPiecesManually Then

    '                For Each sr As DataRow In dsScannedBrcodes.Tables(0).Rows
    '                    If (ur.Cells("BarCode").Value = sr("SerialBarcodeValue")) Then
    '                        ur.Cells("Selected").Value = True
    '                    End If
    '                Next
    '            End If

    '        Next
    '    Catch ex As Exception

    '    End Try
    'End Sub


    'Private Sub tsbAddLines_Click(sender As Object, e As EventArgs) Handles tsbAddLines.Click
    '    Try
    '        Dim objSQL As New clsSqlConn
    '        Dim DS As New DataSet
    '        With objSQL
    '            SQL = "SELECT spilInvNumLines.iInvDetailID AS DetailID, spilInvNumLines.OrderIndex, spilInvNumLines.idInvoiceLines, spilInvNumLines.StockLink, StkItem.Code, StkItem.Description_1," & _
    '                " spilInvNumLines.cDescription As ItemDescription, spilInvNumLines.fQuantity, spilInvNumLines.fThickness As Thickness, spilInvNumLines.iHeight As Height, spilInvNumLines.iWidth As Width," & _
    '                " spilInvNumLines.fVolume As Volume, spilInvNumLines.fQty_Delivered As DeliveredFinishedItems, spilInvNumLines.LineType, spilInvNumLines.Delivery_Status, spilInvNumLines.ItemType, spilInvNumLines.ReBatchQty, " & _
    '                " spilInvNumLines.LineType As LineTypeID, CAST(0 AS bit) As 'Select' FROM spilInvNumLines INNER JOIN StkItem ON spilInvNumLines.StockLink = StkItem.StockLink LEFT  JOIN spilRunnSheetDetailLines ON spilInvNumLines.iInvDetailID = spilRunnSheetDetailLines.iInvDetailID" & _
    '                " WHERE spilInvNumLines.OrderIndex =" & UGSOList.Selected.Rows(0).Cells("OrderIndex").Value & " And (ItemType = 2 Or (ItemType = 1 And M_NO = 0) Or (ItemType = 3 And M_NO = 0) Or (ItemType = 4 And M_NO = 0)) AND spilRunnSheetDetailLines.iInvDetailID is NUll" & _
    '                " ORDER BY spilInvNumLines.idInvoiceLines ASC"

    '            DS = objSQL.GET_DATA_SQL(SQL)


    '            Dim frm As New frmReCutLines(DS, 0)
    '            frm.iRunningSheet = True
    '            Dim result As DialogResult = frm.ShowDialog()
    '            If result = Windows.Forms.DialogResult.OK Then
    '                Dim drSOLine As DataRow
    '                For Each drSOLine In DS.Tables(0).Rows
    '                    If drSOLine("Select") Then
    '                        ugRow1 = UGSOLines.DisplayLayout.Bands(0).AddNew()

    '                        ugRow1.Cells("OrderIndex").Value = drSOLine("OrderIndex")
    '                        ugRow1.Cells("iInvDetailID").Value = drSOLine("DetailID")
    '                        ugRow1.Cells("LineNo").Value = drSOLine("idInvoiceLines")
    '                        ugRow1.Cells("Description").Value = drSOLine("ItemDescription")
    '                        ugRow1.Cells("Thickness").Value = drSOLine("Thickness")
    '                        ugRow1.Cells("Height").Value = drSOLine("Height")
    '                        ugRow1.Cells("Width").Value = drSOLine("Width")
    '                        ugRow1.Cells("Size").Value = drSOLine("Height") & " X " & drSOLine("Width")

    '                        ugRow1.Cells("OrderQty").Value = drSOLine("fQuantity")
    '                        ugRow1.Cells("PrevDelQty").Value = drSOLine("DeliveredFinishedItems")
    '                        ugRow1.Cells("RecutQty").Value = drSOLine("ReBatchQty")

    '                        ugRow1.Cells("ThisDelQty").Activation = Activation.AllowEdit
    '                        If RunnSheetScanningOption <> RunnSheetScanOptions.ScanPiecesManually Then
    '                            If drSOLine("fQuantity") - drSOLine("DeliveredFinishedItems") > 0 Then
    '                                ugRow1.Cells("ThisDelQty").Value = drSOLine("fQuantity") - drSOLine("DeliveredFinishedItems") - drSOLine("ReBatchQty")

    '                            Else
    '                                ugRow1.Cells("ThisDelQty").Value = 0
    '                            End If
    '                        Else
    '                            ''''''
    '                            ugRow1.Cells("ThisDelQty").Value = GetManuallyScanLinePieces(iOrderIndex, drSOLine("DetailID"))
    '                        End If

    '                        ugRow1.Cells("GlassWeight").Value = GetGlassWeight(ugRow1.Cells("ThisDelQty").Value, drSOLine)

    '                        ugRow1.Cells("LineType").Value = drSOLine("LineType")
    '                        ugRow1.Cells("LineTypeID").Value = drSOLine("LineTypeID")
    '                        ugRow1.Cells("Barcodes").Hidden = True


    '                        'ugRow.Cells("BackOrder").Value = 0

    '                        'If drSOLine("LineType") = 1 Then
    '                        '    ugRow1.Cells("LineType").Value = "Re-Cut"
    '                        '    ugRow1.CellAppearance.ForeColor = Color.Red
    '                        'ElseIf drSOLine("LineType") = 2 Then
    '                        '    ugRow1.Cells("LineType").Value = "NCR"
    '                        '    ugRow1.CellAppearance.ForeColor = Color.Magenta
    '                        'Else
    '                        '    ugRow1.Cells("LineType").Value = "Normal"
    '                        'End If
    '                        'GetBarcodes(ugRow1, iOrderIndex, drSOLine("DetailID"), boolsFound)

    '                        If (ugRow1.Cells("ThisDelQty").Value + ugRow1.Cells("PrevDelQty").Value + ugRow1.Cells("RecutQty").Value) < ugRow1.Cells("OrderQty").Value Then
    '                            ugRow1.CellAppearance.ForeColor = Color.FromArgb(130, 7, 7)
    '                            ugRow1.Cells("ThisDelQty").Appearance.BackColor = Color.Yellow
    '                        Else
    '                            ugRow1.CellAppearance.ForeColor = Color.Green
    '                        End If

    '                        If drSOLine("LineTypeID") = LineState.ReBatched Then
    '                            ugRow1.CellAppearance.ForeColor = Color.Red

    '                            'ElseIf drSOLine("LineTypeID") = LineState.NCR Then
    '                            '    ugRow1.CellAppearance.ForeColor = Color.Magenta
    '                        End If

    '                        FillBackOrdQty()
    '                    End If

    '                Next
    '            End If
    '        End With
    '    Catch ex As Exception

    '    End Try
    'End Sub

    Private Sub UGSOList_KeyUp(sender As Object, e As KeyEventArgs) Handles UGSOList.KeyUp
        Dim ugSORow As UltraGridRow = UGSOList.ActiveRow
        lblOrderNum.Text = ugSORow.Cells("OrderNum").Value.ToString()
        For Each ugRow As UltraGridRow In UGSOLines.Rows
            If ugRow.Cells("OrderIndex").Value = ugSORow.Cells("OrderIndex").Value Then
                ugRow.Hidden = False
            Else
                ugRow.Hidden = True
            End If
        Next
    End Sub

    Private Sub UGSOLines_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles UGSOLines.BeforeExitEditMode
        Try
            Dim ugR As UltraGridRow = UGSOLines.ActiveRow

            If (ugR.Cells("ThisDelQty").Value <= ugR.Cells("OrderQty").Value) Then
                If ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value) < 0 Then
                    ugR.Cells("BackOrder").Value = 0
                Else
                    ugR.Cells("BackOrder").Value = ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value + +ugR.Cells("RecutQty").Value)
                End If

                'If ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value) < 0 Then
                '    ugR.Cells("BalanceToDeliver").Value = 0
                'Else
                '    ugR.Cells("BalanceToDeliver").Value = ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value + ugR.Cells("RecutQty").Value)
                'End If

            Else
                MsgBox("Entered This time delivery quantity is higher than the Order quantity. Please enter valid quantity.", MsgBoxStyle.Critical, "Error in This time delivery")
            End If

            CaculateThisTimeDelivery()

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "UGSOLines_BeforeExitEditMode")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        End Try
    End Sub

    Private Sub utmSaveMySettings_ToolClick(sender As Object, e As UltraWinToolbars.ToolClickEventArgs) Handles utmSaveMySettings.ToolClick
        Dim clsSalesOrderObj As New clsSalesOrder
        Dim filePathCol As New Collection
        Try
            If e.Tool.Key = "reset" Then
                If File.Exists(strAppPath & "\" & strUserName & "RunningSheetDetails.xml") = True Then
                    My.Computer.FileSystem.DeleteFile(strAppPath & "\" & strUserName & "RunningSheetDetails.xml")
                End If
            ElseIf e.Tool.Key = "ResetLines" Then
                If File.Exists(strAppPath & "\" & strUserName & "RunningSheetSODetails.xml") = True Then
                    My.Computer.FileSystem.DeleteFile(strAppPath & "\" & strUserName & "RunningSheetSODetails.xml")
                End If
            End If
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "utmSaveMySettings_ToolClick")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        End Try
    End Sub

    Private Sub UGSOLines_KeyUp(sender As Object, e As KeyEventArgs) Handles UGSOLines.KeyUp
        Try
            Dim ugR As UltraGridRow = UGSOLines.ActiveRow

            If (ugR.Cells("ThisDelQty").Value <= ugR.Cells("OrderQty").Value) Then
                If ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value) < 0 Then
                    ugR.Cells("BackOrder").Value = 0
                Else
                    ugR.Cells("BackOrder").Value = ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value + +ugR.Cells("RecutQty").Value)
                End If

                'If ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value) < 0 Then
                '    ugR.Cells("BalanceToDeliver").Value = 0
                'Else
                '    ugR.Cells("BalanceToDeliver").Value = ugR.Cells("OrderQty").Value - (ugR.Cells("PrevDelQty").Value + ugR.Cells("ThisDelQty").Value + ugR.Cells("RecutQty").Value)
                'End If

            Else
                MsgBox("Entered This time delivery quantity is higher than the Order quantity. Please enter valid quantity.", MsgBoxStyle.Critical, "Error in This time delivery")
            End If

            CaculateThisTimeDelivery()

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "UGSOLines_KeyUp")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        End Try
    End Sub

    Private Sub cmdSaveSettings_Click(sender As Object, e As EventArgs) Handles cmdSaveSettings.Click
        UGSOList.DisplayLayout.SaveAsXml(strAppPath & "\" & strUserName & "RunningSheetDetails.xml")
    End Sub

    Private Sub CmdSaveMySettingSOLines_Click(sender As Object, e As EventArgs) Handles CmdSaveMySettingSOLines.Click
        UGSOLines.DisplayLayout.SaveAsXml(strAppPath & "\" & strUserName & "RunningSheetSODetails.xml")
    End Sub

    Private Sub CaculateThisTimeDelivery()
        Try
            Dim iThisTimeQty As Integer = 0

            If (UGSOLines.Rows.Count > 0) Then
                Dim ugR As UltraGridRow
                For Each ugR In UGSOLines.Rows
                    If (ugR.Hidden = False) Then
                        iThisTimeQty += ugR.Cells("ThisDelQty").Value
                    End If

                    If (ugR.Cells("ThisDelQty").Value + ugR.Cells("PrevDelQty").Value + ugR.Cells("RecutQty").Value) < ugR.Cells("OrderQty").Value Then
                        ugR.CellAppearance.ForeColor = Color.FromArgb(130, 7, 7)
                        ugR.Cells("ThisDelQty").Appearance.BackColor = Color.Yellow
                    Else
                        ugR.CellAppearance.ForeColor = Color.Green
                    End If

                Next
            End If

            UGSOList.ActiveRow.Cells("ThisTimeDelivery").Value = iThisTimeQty
            UGSOList.ActiveRow.Cells("BalanceToDeliver").Value = CInt(UGSOList.ActiveRow.Cells("TotalGlassPanels").Value) - CInt(UGSOList.ActiveRow.Cells("ThisTimeDelivery").Value)

            If (CInt(UGSOList.ActiveRow.Cells("ThisTimeDelivery").Value) + CInt(UGSOList.ActiveRow.Cells("DeliveredSoFar").Value)) < UGSOList.ActiveRow.Cells("TotalGlassPanels").Value Then
                UGSOList.ActiveRow.Appearance.ForeColor = Color.FromArgb(130, 7, 7)
            Else
                UGSOList.ActiveRow.Appearance.ForeColor = Color.FromArgb(4, 73, 6)
            End If
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "SQL Error", "CaculateThisTimeDelivery")
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL Glass")
        End Try
    End Sub

    Function SaveUsingParaDetails(ByRef isNewRecord As Boolean, ByRef clsSqlConnObj As clsSqlConn) As Integer
        Dim collspPara As New Collection
        Dim colPara As New spParameters
        Dim newSQLQuery As String = ""
        Try
            colPara.ParaName = "@RunnNo"
            colPara.ParaValue = If(IsNothing(txtRunnNumber.Value) = False, txtRunnNumber.Value, 0)
            collspPara.Add(colPara)

            colPara.ParaName = "@AreaID"
            colPara.ParaValue = If(IsNothing(cboArea.Value) = False, cboArea.Value, 0)
            collspPara.Add(colPara)

            colPara.ParaName = "@RunnDate"
            colPara.ParaValue = If(IsNothing(txtRunningDate.Value) = False, txtRunningDate.Value, "01/01/1900 00:00:00 AM")
            collspPara.Add(colPara)

            colPara.ParaName = "@RunnTime"
            colPara.ParaValue = If(IsNothing(txtRunnTime.Value) = False, txtRunnTime.Value, "01/01/1900 00:00:00 AM")
            collspPara.Add(colPara)

            colPara.ParaName = "@VehRegNo"
            colPara.ParaValue = If(IsNothing(txtVehRegNo) = False, txtVehRegNo.Text.Replace("'", ""), "")
            collspPara.Add(colPara)

            colPara.ParaName = "@DrivName"
            colPara.ParaValue = If(IsNothing(txtDrivName) = False, txtDrivName.Text.Replace("'", ""), "")
            collspPara.Add(colPara)

            colPara.ParaName = "@Reference"
            colPara.ParaValue = If(IsNothing(txtReference) = False, txtRunnNumber.Text.Replace("'", ""), "")
            collspPara.Add(colPara)

            colPara.ParaName = "@TelNo"
            colPara.ParaValue = If(IsNothing(txtTeleNo) = False, txtTeleNo.Text.Replace("'", ""), "")
            collspPara.Add(colPara)

            colPara.ParaName = "@Notes"
            colPara.ParaValue = If(IsNothing(txtNotes) = False, txtNotes.Text.Replace("'", ""), "")
            collspPara.Add(colPara)

            colPara.ParaName = "@DocPrinted"
            colPara.ParaValue = 0
            collspPara.Add(colPara)

            colPara.ParaName = "@EnteredBy"
            colPara.ParaValue = If(IsNothing(strUserName) = False, strUserName, "")
            collspPara.Add(colPara)

            colPara.ParaName = "@EnteredDateTime"
            colPara.ParaValue = Now
            collspPara.Add(colPara)

            colPara.ParaName = "@Status"
            colPara.ParaValue = GlassReceiptState.UnProcessed
            collspPara.Add(colPara)

            colPara.ParaName = "@FacilityID"
            colPara.ParaValue = If(IsNothing(cmbFacility.Value) = False, cmbFacility.Value, 1)
            collspPara.Add(colPara)

            colPara.ParaName = "@Duration"
            colPara.ParaValue = If(IsNothing(cmbDuration.Value) = False, cmbDuration.Value, 0)
            collspPara.Add(colPara)

            colPara.ParaName = "@directionURL"
            colPara.ParaValue = If(IsNothing(directionURL) = False, directionURL, "")
            collspPara.Add(colPara)

            colPara.ParaName = "@googleMapImagePath"
            colPara.ParaValue = If(IsNothing(GetImageLocationPath(txtRunnNumber.Value, False)) = False, GetImageLocationPath(txtRunnNumber.Value, False), "")
            collspPara.Add(colPara)

            If isNewRecord = True Then
                newSQLQuery = "set dateformat dmy Insert into spilRunnSheetHeader " &
            "(RunnNo,  AreaID,  RunnDate,  RunnTime,  VehRegNo,  DrivName,  Reference,  TelNo,  Notes, " &
            "DocPrinted,  EnteredBy,  EnteredDateTime,  Status,  FacilityID,  Duration, directionURL) " &
            "values(@RunnNo,  @AreaID,  @RunnDate,  @RunnTime,  @VehRegNo,  @DrivName,  @Reference,  @TelNo, " &
            "@Notes,  @DocPrinted,  @EnteredBy,  @EnteredDateTime,  @Status,  @FacilityID,  @Duration, @directionURL)"
            Else
                newSQLQuery = "set dateformat dmy update spilRunnSheetHeader " &
            "set AreaID= @AreaID, RunnDate= @RunnDate, RunnTime= @RunnTime, VehRegNo= @VehRegNo, DrivName= @DrivName" &
            ", Reference= @Reference, TelNo= @TelNo, Notes= @Notes, EnteredBy= @EnteredBy, EnteredDateTime= @EnteredDateTime" &
            ", FacilityID= @FacilityID, Duration= @Duration, directionURL=@directionURL where RunnNo=@RunnNo"
            End If

            Return clsSqlConnObj.EXE_SQL_Trans_Para_Return(newSQLQuery, collspPara)

        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical, "warning", GetCurrentMethod.Name)
            Return 0
        End Try
    End Function

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles pbGoogleMap.Click
        'GetMapURL()
    End Sub

    Private Sub UGSOList_AfterRowsDeleted(sender As Object, e As EventArgs) Handles UGSOList.AfterRowsDeleted
        btnRefreshGoogleMap.Visible = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnRefreshGoogleMap.Click
        btnRefreshGoogleMap.Visible = False
        GetMapURL()
    End Sub

    Private Sub UGSOList_AfterRowInsert(sender As Object, e As RowEventArgs) Handles UGSOList.AfterRowInsert
        btnRefreshGoogleMap.Visible = True

    End Sub
End Class