Imports Microsoft.VisualBasic
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace TestExcelRTDServer
	<Guid("B6AF4673-200B-413c-8536-1F778AC14DE1"), ProgId("My.Sample.RtdServer"), ComVisible(True)> _
	Public Class RtdServer
		Implements IRtdServer
		Private m_callback As IRTDUpdateEvent
		Private m_timer As Timer
		Private m_topics As Dictionary(Of Integer, String)
		Private Shared random As New Random()

		Public Function ServerStart(ByVal callback As IRTDUpdateEvent) As Integer Implements IRtdServer.ServerStart
			m_callback = callback

			m_timer = New Timer()
			AddHandler m_timer.Tick, AddressOf TimerEventHandler
			m_timer.Interval = 2000

			m_topics = New Dictionary(Of Integer, String)()

			Return 1
		End Function

		Public Sub ServerTerminate() Implements IRtdServer.ServerTerminate
			If Nothing IsNot m_timer Then
				m_timer.Dispose()
				m_timer = Nothing
			End If
		End Sub

		Public Function ConnectData(ByVal topicId As Integer, ByRef strings As Array, ByRef newValues As Boolean) As Object Implements IRtdServer.ConnectData
			If 1 <> strings.Length Then
				Return "Exactly one parameter is required"
			End If

			Dim value As String = strings.GetValue(0).ToString()

			m_topics(topicId) = value
			m_timer.Start()
			Return GetNextValue(value)
		End Function

		Public Sub DisconnectData(ByVal topicId As Integer) Implements IRtdServer.DisconnectData
			m_topics.Remove(topicId)
		End Sub

		Public Function RefreshData(ByRef topicCount As Integer) As Array Implements IRtdServer.RefreshData
			Dim data(1, m_topics.Count - 1) As Object

			Dim index As Integer = 0

			For Each topicId As Integer In m_topics.Keys
				data(0, index) = topicId
				data(1, index) = GetNextValue(m_topics(topicId))

				index += 1
			Next topicId

			topicCount = m_topics.Count

			m_timer.Start()
			Return data
		End Function

		Public Function Heartbeat() As Integer Implements IRtdServer.Heartbeat
			Return 1
		End Function

		Private Sub TimerEventHandler(ByVal sender As Object, ByVal args As EventArgs)
			m_timer.Stop()
			m_callback.UpdateNotify()
		End Sub

		Private Shared Function GetNextValue(ByVal value As String) As Double
			Dim quote As Double
			Select Case value
				Case "MSFT"
					quote = 40
				Case "FB"
					quote = 60
				Case "YHOO"
					quote = 36
				Case Else
					quote = Double.Parse(value)
			End Select
			Return quote + (random.NextDouble() * 10.0 - 5.0)
		End Function
	End Class
End Namespace

Namespace Microsoft.Office.Interop.Excel
	<Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8"), InterfaceType(ComInterfaceType.InterfaceIsDual), ComImport(), TypeIdentifier, ComVisible(True)> _
	Public Interface IRTDUpdateEvent
		Sub UpdateNotify()

'        int HeartbeatInterval { get; set; }

'        void Disconnect();
	End Interface
	<Guid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8"), InterfaceType(ComInterfaceType.InterfaceIsDual), ComImport(), TypeIdentifier, ComVisible(True)> _
	Public Interface IRtdServer
		<DispId(10)> _
		Function ServerStart(ByVal callback As IRTDUpdateEvent) As Integer
		<DispId(11)> _
		Function ConnectData(ByVal topicId As Integer, <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType := VarEnum.VT_VARIANT)> ByRef strings As Array, ByRef newValues As Boolean) As Object

		<DispId(12)> _
		Function RefreshData(ByRef topicCount As Integer) As <MarshalAs(UnmanagedType.SafeArray, SafeArraySubType := VarEnum.VT_VARIANT)> Array

		<DispId(13)> _
		Sub DisconnectData(ByVal topicId As Integer)

		<DispId(14)> _
		Function Heartbeat() As Integer

		<DispId(15)> _
		Sub ServerTerminate()
	End Interface
End Namespace
