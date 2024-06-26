<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613681/16.2.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E5204)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/TestRTDClient/Form1.cs) (VB: [Form1.vb](./VB/TestRTDClient/Form1.vb))
* [RtdServer.cs](./CS/TestRTDServer/RtdServer.cs) (VB: [RtdServer.vb](./VB/TestRTDServer/RtdServer.vb))
<!-- default file list end -->
# How to implement real-time data server and use RTD worksheet function


<p>This example illustrates how to set up and use RTD worksheet function. The Real-Time Data (RTD) function enables you to retrieve data in real time from a COM Automation server. The function result is updated when new data becomes available from the server.<br> The example contains the code of a simple RTD server that provides random data. You can revise the code or modify the server's GetValue method to implement complex scenarios.</p>
<p>To run the project, start the Microsoft Visual Studio as administrator. Elevated permissions are required to register COM server after it is built. Build the project, then run it.<br> The following animated image illustrates the resulting application.</p>
<p><img src="https://raw.githubusercontent.com/DevExpress-Examples/how-to-implement-real-time-data-server-and-use-rtd-worksheet-function-e5204/16.2.4+/media/f5b2d302-e469-4bd5-af24-6b33570f7c0f.png"><br><br>In version 16.2.4 and higher you can use the <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetDocumentOptions_RealTimeDatatopic">SpreadsheetControl.Options.RealTimeData</a> property to specify whether updates are manual or by timer, and set the time interval between updates.</p>

<br/>


<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-implement-real-time-data-server-and-use-rtd-worksheet-function&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=how-to-implement-real-time-data-server-and-use-rtd-worksheet-function&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
