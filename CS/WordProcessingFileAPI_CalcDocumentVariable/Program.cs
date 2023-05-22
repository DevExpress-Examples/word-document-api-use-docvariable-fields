using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordProcessingFileAPI_CalcDocumentVariable
{
    class Program
    {
        static void Main(string[] args)
        {
            RichEditDocumentServer wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument("Docs\\invitation.docx", DocumentFormat.OpenXml);

            Document document = wordProcessor.Document;
            //Lock the first field to prevent updates
            document.Fields[0].Locked = true;

            // Handle the CalculateDocumentVariable event
            document.CalculateDocumentVariable += Document_CalculateDocumentVariable;

            // Adjust mail-merge options
            MailMergeOptions myMergeOptions = document.CreateMailMergeOptions();
            myMergeOptions.MergeMode = MergeMode.NewSection;
            myMergeOptions.DataSource = new SampleData();

            // Handle mail-merge events
            wordProcessor.MailMergeRecordStarted += wordProcessor_MailMergeRecordStarted;
            wordProcessor.MailMergeRecordFinished += wordProcessor_MailMergeRecordFinished;

            // Mail-merge the document
            document.MailMerge(myMergeOptions, "Result.docx", DocumentFormat.OpenXml);

            Process.Start("Result.docx");
        }

        private static void Document_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
        {
            if (e.Arguments.Count > 0)
            {
                string location = e.Arguments[0].Value.ToString();
                if ((location.Trim() == String.Empty) || (location.Contains("<")))
                {
                    e.Value = " ";
                    e.Handled = true;
                    return;
                }
                switch (e.VariableName)
                {
                    case "Weather":
                        Conditions conditions = new Conditions();
                        conditions = Weather.GetCurrentConditions(location);
                        e.Value = String.Format("Weather for {0}: \nConditions: {1}\nTemperature (C) :{2}\nHumidity: {3}\nWind: {4}\n",
                            location, conditions.Condition, conditions.TempC, conditions.Humidity, conditions.Wind);
                        break;
                    case "LOCATION":
                        if (location == "DO NOT CHANGE!") e.Value = DocVariableValue.Current;
                        break;
                    default:
                        e.Value = "LOCKED FIELD UPDATED";
                        break;
                }
            }
            else
            {
                e.Value = "LOCKED FIELD UPDATED";
            }
            e.Handled = true;
        }

        private static void wordProcessor_MailMergeRecordStarted(object sender, MailMergeRecordStartedEventArgs e)
            {
                DocumentRange insertedRange = e.RecordDocument.InsertText(e.RecordDocument.Range.Start, String.Format("Created on {0:G}\n\n", DateTime.Now));
                CharacterProperties cp = e.RecordDocument.BeginUpdateCharacters(insertedRange);
                cp.FontSize = 8;
                cp.ForeColor = Color.Red;
                e.RecordDocument.EndUpdateCharacters(cp);
            }

            private static void wordProcessor_MailMergeRecordFinished(object sender, MailMergeRecordFinishedEventArgs e)
            {
                e.RecordDocument.AppendDocumentContent("Docs\\bungalow.docx", DocumentFormat.OpenXml);
            }
    }
}
