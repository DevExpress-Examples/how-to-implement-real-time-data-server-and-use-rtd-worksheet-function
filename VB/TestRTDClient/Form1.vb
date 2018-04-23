Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace TestRTDClient
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

			spreadsheetControl1.LoadDocument("Portfolio.xlsx")
		End Sub
	End Class
End Namespace
