Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Namespace TestRTDClient

    Public Partial Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            spreadsheetControl1.LoadDocument("Portfolio.xlsx")
        End Sub
    End Class
End Namespace
