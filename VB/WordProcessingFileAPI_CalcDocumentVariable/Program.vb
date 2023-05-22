Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Diagnostics
Imports System.Drawing

Namespace WordProcessingFileAPI_CalcDocumentVariable

    Friend Class Program

        Shared Sub Main(ByVal args As String())
            Dim wordProcessor As RichEditDocumentServer = New RichEditDocumentServer()
            wordProcessor.LoadDocument("Docs\invitation.docx", DocumentFormat.OpenXml)

            Dim document As Document = wordProcessor.Document
            ' Lock the first field to prevent updates
            document.Fields(0).Locked = True

            ' Handle the CalculateDocumentVariable event
            AddHandler document.CalculateDocumentVariable, AddressOf Document_CalculateDocumentVariable

            ' Adjust mail-merge options
            Dim myMergeOptions As MailMergeOptions = document.CreateMailMergeOptions()
            myMergeOptions.MergeMode = MergeMode.NewSection
            myMergeOptions.DataSource = New SampleData()

            ' Handle mail-merge events
            AddHandler wordProcessor.MailMergeRecordStarted, AddressOf wordProcessor_MailMergeRecordStarted
            AddHandler wordProcessor.MailMergeRecordFinished, AddressOf wordProcessor_MailMergeRecordFinished

            ' Mail-merge the document
            document.MailMerge(myMergeOptions, "Result.docx", DocumentFormat.OpenXml)
            Call Process.Start("Result.docx")
        End Sub

        Private Shared Sub Document_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
            If e.Arguments.Count > 0 Then
                Dim location As String = e.Arguments(0).Value.ToString()
                If(Equals(location.Trim(), String.Empty)) OrElse location.Contains("<") Then
                    e.Value = " "
                    e.Handled = True
                    Return
                End If

                Select Case e.VariableName
                    Case "Weather"
                        Dim conditions As Conditions = New Conditions()
                        conditions = Weather.GetCurrentConditions(location)
                        e.Value = String.Format("Weather for {0}: " & Microsoft.VisualBasic.Constants.vbLf & "Conditions: {1}" & Microsoft.VisualBasic.Constants.vbLf & "Temperature (C) :{2}" & Microsoft.VisualBasic.Constants.vbLf & "Humidity: {3}" & Microsoft.VisualBasic.Constants.vbLf & "Wind: {4}" & Microsoft.VisualBasic.Constants.vbLf, location, conditions.Condition, conditions.TempC, conditions.Humidity, conditions.Wind)
                    Case "LOCATION"
                        If Equals(location, "DO NOT CHANGE!") Then e.Value = DocVariableValue.Current
                    Case Else
                        e.Value = "LOCKED FIELD UPDATED"
                End Select
            Else
                e.Value = "LOCKED FIELD UPDATED"
            End If

            e.Handled = True
        End Sub

        Private Shared Sub wordProcessor_MailMergeRecordStarted(ByVal sender As Object, ByVal e As MailMergeRecordStartedEventArgs)
            Dim insertedRange As DocumentRange = e.RecordDocument.InsertText(e.RecordDocument.Range.Start, String.Format("Created on {0:G}" & Microsoft.VisualBasic.Constants.vbLf & Microsoft.VisualBasic.Constants.vbLf, Date.Now))
            Dim cp As CharacterProperties = e.RecordDocument.BeginUpdateCharacters(insertedRange)
            cp.FontSize = 8
            cp.ForeColor = Color.Red
            e.RecordDocument.EndUpdateCharacters(cp)
        End Sub

        Private Shared Sub wordProcessor_MailMergeRecordFinished(ByVal sender As Object, ByVal e As MailMergeRecordFinishedEventArgs)
            e.RecordDocument.AppendDocumentContent("Docs\bungalow.docx", DocumentFormat.OpenXml)
        End Sub
    End Class
End Namespace
