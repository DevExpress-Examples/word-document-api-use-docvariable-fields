Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace WordProcessingFileAPI_CalcDocumentVariable
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Dim srv As New RichEditDocumentServer()
			srv.LoadDocument("Docs\invitation.docx")

			srv.Document.Fields(0).Locked = True

			srv.Options.MailMerge.DataSource = New SampleData()
			AddHandler srv.Document.CalculateDocumentVariable, AddressOf Document_CalculateDocumentVariable
			Dim myMergeOptions As MailMergeOptions = srv.Document.CreateMailMergeOptions()

			AddHandler srv.MailMergeRecordStarted, AddressOf srv_MailMergeRecordStarted
			AddHandler srv.MailMergeRecordFinished, AddressOf srv_MailMergeRecordFinished

			myMergeOptions.MergeMode = MergeMode.NewSection
			srv.Document.MailMerge(myMergeOptions, "Result.docx", DocumentFormat.OpenXml)

			Process.Start("Result.docx")
		End Sub

		Private Shared Sub Document_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
			If e.Arguments.Count > 0 Then
				Dim location As String = e.Arguments(0).Value.ToString()
				If (location.Trim() = String.Empty) OrElse (location.Contains("<")) Then
					e.Value = " "
					e.Handled = True
					Return
				End If
				Select Case e.VariableName
					Case "Weather"
						Dim conditions As New Conditions()
						conditions = Weather.GetCurrentConditions(location)
						e.Value = String.Format("Weather for {0}: " & vbLf & "Conditions: {1}" & vbLf & "Temperature (C) :{2}" & vbLf & "Humidity: {3}" & vbLf & "Wind: {4}" & vbLf, location, conditions.Condition, conditions.TempC, conditions.Humidity, conditions.Wind)
					Case "LOCATION"
						If location = "DO NOT CHANGE!" Then
							e.Value = DocVariableValue.Current
						End If
					Case Else
						e.Value = "LOCKED FIELD UPDATED"
				End Select
			Else
				e.Value = "LOCKED FIELD UPDATED"
			End If
			e.Handled = True
		End Sub

		Private Shared Sub srv_MailMergeRecordStarted(ByVal sender As Object, ByVal e As MailMergeRecordStartedEventArgs)
				Dim _range As DocumentRange = e.RecordDocument.InsertText(e.RecordDocument.Range.Start, String.Format("Created on {0:G}" & vbLf & vbLf, DateTime.Now))
				Dim cp As CharacterProperties = e.RecordDocument.BeginUpdateCharacters(_range)
				cp.FontSize = 8
				cp.ForeColor = Color.Red
				cp.Hidden = True
				e.RecordDocument.EndUpdateCharacters(cp)
		End Sub

			Private Shared Sub srv_MailMergeRecordFinished(ByVal sender As Object, ByVal e As MailMergeRecordFinishedEventArgs)
				e.RecordDocument.AppendDocumentContent("Docs\bungalow.docx", DocumentFormat.OpenXml)
			End Sub
	End Class
End Namespace
