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
        Private m_topics As Dictionary(Of Integer, TopicData)
        Private companies As Dictionary(Of String, CompanyData)

        Public Sub New()
            Me.companies = New Dictionary(Of String, CompanyData)()
            Me.companies.Add("MSFT", New CompanyData("MSFT", 40, 176))
            Me.companies.Add("FB", New CompanyData("FB", 60, 210))
            Me.companies.Add("YHOO", New CompanyData("YHOO", 36, 54))
            Me.companies.Add("NOK", New CompanyData("NOK", 50, 100))
            UpdatePrices()
        End Sub

        Public Function ServerStart(ByVal callback As IRTDUpdateEvent) As Integer Implements IRtdServer.ServerStart
            m_callback = callback

            m_timer = New Timer()
            AddHandler m_timer.Tick, AddressOf TimerEventHandler
            m_timer.Interval = 500

            m_topics = New Dictionary(Of Integer, TopicData)()

            Return 1
        End Function

        Public Sub ServerTerminate() Implements IRtdServer.ServerTerminate
            If Nothing IsNot m_timer Then
                m_timer.Dispose()
                m_timer = Nothing
            End If
        End Sub

        Public Function ConnectData(ByVal topicId As Integer, ByRef strings As Array, ByRef newValues As Boolean) As Object Implements IRtdServer.ConnectData
            If 2 <> strings.Length Then
                Return "Two parameters are required"
            End If
            Dim symbol As String = strings.GetValue(0).ToString()
            Dim type As String = strings.GetValue(1).ToString()

            Dim data As New TopicData(symbol, type)

            m_topics(topicId) = data
            m_timer.Start()
            Return GetNextValue(data)
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
            UpdatePrices()
            m_callback.UpdateNotify()
        End Sub

        Private Function GetNextValue(ByVal data As TopicData) As Object
            Dim companyData As CompanyData = Nothing
            If companies.TryGetValue(data.Symbol, companyData) Then
                Select Case data.Type
                    Case "PRICE"
                        Return companyData.Price
                    Case "CHANGE"
                        Return companyData.Change
                    Case "SHARES"
                        Return companyData.Shares
                End Select
            End If
            Return "#Invalid"
        End Function

        Private Sub UpdatePrices()
            For Each companyData As CompanyData In companies.Values
                companyData.UpdatePrice()
            Next companyData
        End Sub
    End Class
End Namespace

Public Class TopicData
    Public Sub New(ByVal symbol As String, ByVal type As String)
        Me.Symbol = symbol
        Me.Type = type
    End Sub

    Public ReadOnly Property Symbol() As String
    Public ReadOnly Property Type() As String
End Class

Public Class CompanyData
    Private Shared random As New Random()

    Public Sub New(ByVal symbol As String, ByVal quote As Double, ByVal shares As Integer)
        Me.Symbol = symbol
        Me.Quote = quote
        Me.Shares = shares
    End Sub

    Private ReadOnly Property Quote() As Double
    Public ReadOnly Property Symbol() As String
    Public ReadOnly Property Shares() As Integer
    Private privatePrice As Double
    Public Property Price() As Double
        Get
            Return privatePrice
        End Get
        Private Set(ByVal value As Double)
            privatePrice = value
        End Set
    End Property
    Private privateChange As Double
    Public Property Change() As Double
        Get
            Return privateChange
        End Get
        Private Set(ByVal value As Double)
            privateChange = value
        End Set
    End Property

    Public Sub UpdatePrice()
        Me.Price = Quote + (random.NextDouble() * 10.0 - 5.0)
        Me.Change = (Me.Price - Me.Quote) / Me.Quote
    End Sub
End Class

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
