using DevExpress.XtraSpreadsheet;
using System;

namespace TestRTDClient {
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm {
        public Form1()
        {
            InitializeComponent();

            spreadsheetControl1.LoadDocument("Portfolio.xlsx");
        }

        private void barEditRefreshMode_EditValueChanged(object sender, EventArgs e) {
            RealTimeDataRefreshMode mode = (RealTimeDataRefreshMode)Enum.Parse(typeof(RealTimeDataRefreshMode), barEditRefreshMode.EditValue.ToString());
            spreadsheetControl1.Options.RealTimeData.RefreshMode = mode;
        }

        private void barEditSpinThrottleInterval_EditValueChanged(object sender, EventArgs e) {
            spreadsheetControl1.Options.RealTimeData.ThrottleInterval = Int32.Parse(barEditSpinThrottleInterval.EditValue.ToString());
        }

        private void barbButtonRefreshData_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            spreadsheetControl1.Document.RealTimeData.RefreshData();
        }

        private void barButtonRestartServers_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            spreadsheetControl1.Document.RealTimeData.RestartServers();
        }
    }
}
