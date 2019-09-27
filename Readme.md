<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/TestRTDClient/Form1.cs) (VB: [Form1.vb](./VB/TestRTDClient/Form1.vb))
* [RtdServer.cs](./CS/TestRTDServer/RtdServer.cs) (VB: [RtdServer.vb](./VB/TestRTDServer/RtdServer.vb))
<!-- default file list end -->
# How to implement a real-time data server and use the RTD function

This example demonstrates how to use the RTD function to retrieve data in real time from a COM Automation server.

In this example, we use a custom server that implements the [IRtdServer](https://docs.microsoft.com/en-US/dotnet/api/microsoft.office.interop.excel.irtdserver) interface. Our server provides data for stock prices, number of shares, and price change.

To run the project, start Microsoft Visual Studio as an administrator. Elevated permissions are required to register the COM server after it is built.

The image below shows the resulting application.
![Spreadsheet_RTD](/media/rtd-function.gif)

In v16.2.4 and higher, you can use the [SpreadsheetControl.Options.RealTimeData](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.DocumentOptions.RealTimeData) property to specify whether to update data manually or to use a timer for automatic updates.
