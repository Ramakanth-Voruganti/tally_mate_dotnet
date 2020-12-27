
Imports System.IO

Imports System.Threading
Imports System.Threading.Tasks


Public Class frm_master_det

    Dim oFirmDet As New clsFirmDet

    Dim oMasterDetTlyBridge As clsMasterDet_TallyBridge

    Dim oLedgMasterTly As New clsLedgerMaster_Tally

    Dim oWbookSalesAct As New clsExcelWbook


#Region "dgrid_master_actual"

    Const ColSNo = "SNo"
    Const ColName = "Name"

    Dim WithEvents oDgvMasterAct As clsDataGridView

#End Region


    Dim oPbr As clsProgressbar

    Private Sub SetFocus(ByVal Ctrl_par As Object)

        Try
            Me.ActiveControl = Ctrl_par
        Catch

        End Try

    End Sub


    Private Sub frm_master_det_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Disposed

        Dim mpointer_lcl = Cursor.Current
        Cursor.Current = Cursors.WaitCursor

        '        Try

        If (Not IsNothing(ExcelApp)) Then
            ExcelApp.Quit()
        End If

        oFirmDet.Dispose()

        '        Catch

        '        End Try

end_sub:

        Cursor.Current = mpointer_lcl

    End Sub

    Private Sub frm_master_det_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim mpointer_lcl = Cursor.Current
        Cursor.Current = Cursors.Default

        ' Related to Cross Thread Reference of Controls
        '        Control.CheckForIllegalCrossThreadCalls = False

        oPbr = New clsProgressbar(pbr_frm_import_sales_to_tally)


        With oFirmDet

            ._DbPath = Replace(LCase(Application.StartupPath), LCase("\bin\Debug"), "\",,, CompareMethod.Text)    ' "h:\tally_mate\tally_mate_dotNet"
            ._DbName = "sss_tdf.mdb"

            .OpenMainDbOleConn()

        End With


        tab_import_main.TabIndex = 0

end_sub:
        Cursor.Current = mpointer_lcl

    End Sub

    Private Sub btn_post_xml_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_post_xml.Click

        Dim lst_data As New List(Of clsXmlTag)


        Dim oGrpMaster As New clsGroupMaster_Tally

        With oGrpMaster

            With .XmlDet

                ._tagGrpName.Data = "test-wgl"

                '                MsgBox(.CreateDet)

                '                MsgBox(ParseXmlText(.FindDet("test", "", "9000"), XmlParseByTallyMessage, lst_data))

            End With

        End With


        Dim oLedgMaster As New clsLedgerMaster_Tally

        With oLedgMaster

            With .XmlDet

                ._tagLedgName.Data = txt_name.Text
                ._tagName.Data = ._tagLedgName.Data

                txt_group_Leave(txt_group, EventArgs.Empty)
                ._tagParent.Data = txt_group.Text

                ._tagAddr1.Data = txt_address.Text

                MsgBox(.CreateDet("test", "", "9000"))

                '                MsgBox(ParseXmlText(.FindDet("test", "", "9000"), XmlParseByTallyMessage, lst_data))

            End With

        End With


end_sub:

    End Sub

    Private Sub dgrid_master_actual_MouseWheel(ByVal sender As Object,
                                              ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgrid_master_actual.MouseWheel

        Dim intMove As Integer = dgrid_master_actual.FirstDisplayedScrollingRowIndex - e.Delta / 120

        If (intMove >= 0) And (intMove <= dgrid_master_actual.Rows.Count) Then
            dgrid_master_actual.FirstDisplayedScrollingRowIndex = intMove
        End If

    End Sub

    Private Sub txt_group_Leave(sender As Object, e As EventArgs) Handles txt_group.Leave

        With txt_group

            If (Len(.Text) = 0) Then
                .Text = "Capital Account"
            End If

        End With

    End Sub

End Class