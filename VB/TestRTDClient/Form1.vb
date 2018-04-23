Imports DevExpress.XtraSpreadsheet

Namespace TestRTDClient
    Partial Public Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Public Sub New()
            InitializeComponent()

            spreadsheetControl1.LoadDocument("Portfolio.xlsx")
        End Sub

        Private Sub barEditRefreshMode_EditValueChanged(ByVal sender As Object, ByVal e As EventArgs) Handles barEditRefreshMode.EditValueChanged
            Dim mode As RealTimeDataRefreshMode = DirectCast(System.Enum.Parse(GetType(RealTimeDataRefreshMode), barEditRefreshMode.EditValue.ToString()), RealTimeDataRefreshMode)
            spreadsheetControl1.Options.RealTimeData.RefreshMode = mode
        End Sub

        Private Sub barEditSpinThrottleInterval_EditValueChanged(ByVal sender As Object, ByVal e As EventArgs) Handles barEditSpinThrottleInterval.EditValueChanged
            spreadsheetControl1.Options.RealTimeData.ThrottleInterval = Int32.Parse(barEditSpinThrottleInterval.EditValue.ToString())
        End Sub

        Private Sub barbButtonRefreshData_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barbButtonRefreshData.ItemClick
            spreadsheetControl1.Document.RealTimeData.RefreshData()
        End Sub

        Private Sub barButtonRestartServers_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles barButtonRestartServers.ItemClick
            spreadsheetControl1.Document.RealTimeData.RestartServers()
        End Sub
    End Class
End Namespace
