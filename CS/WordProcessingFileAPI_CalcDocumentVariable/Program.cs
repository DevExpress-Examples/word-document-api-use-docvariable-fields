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
            RichEditDocumentServer srv = new RichEditDocumentServer();
            srv.LoadDocument("Docs\\invitation.docx");

            srv.Document.Fields[0].Locked = true;

            srv.Options.MailMerge.DataSource = new SampleData();
            srv.Document.CalculateDocumentVariable += Document_CalculateDocumentVariable;
            MailMergeOptions myMergeOptions = srv.Document.CreateMailMergeOptions();

            srv.MailMergeRecordStarted += srv_MailMergeRecordStarted;
            srv.MailMergeRecordFinished += srv_MailMergeRecordFinished;

            myMergeOptions.MergeMode = MergeMode.NewSection;
            srv.Document.MailMerge(myMergeOptions, "Result.docx", DocumentFormat.OpenXml);

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

        private static void srv_MailMergeRecordStarted(object sender, MailMergeRecordStartedEventArgs e)
            {
                DocumentRange _range = e.RecordDocument.InsertText(e.RecordDocument.Range.Start, String.Format("Created on {0:G}\n\n", DateTime.Now));
                CharacterProperties cp = e.RecordDocument.BeginUpdateCharacters(_range);
                cp.FontSize = 8;
                cp.ForeColor = Color.Red;
                cp.Hidden = true;
                e.RecordDocument.EndUpdateCharacters(cp);
            }

            private static void srv_MailMergeRecordFinished(object sender, MailMergeRecordFinishedEventArgs e)
            {
                e.RecordDocument.AppendDocumentContent("Docs\\bungalow.docx", DocumentFormat.OpenXml);
            }
    }
}
