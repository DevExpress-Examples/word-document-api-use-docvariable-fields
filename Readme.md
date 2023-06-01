<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/155868028/18.2.2%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830501)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Word Processing Document API - How to Use DOCVARIABLE Fields in a Document

This example illustrates the use of a **DOCVARIABLE** field to provide additional information which is dependent on the value of a merged field. This technique is implemented so each merged document contains a weather report for a location that corresponds to the current data record.

> **Note**
>
> We do not provide code for retrieving weather information. You can implement a custom weather information provider.

## Implementation Details

A MERGEFIELD field defines a location. The field is included as an argument in the DOCVARIABLE field. When the DOCVARIABLE field is updated, the [Document.CalculateDocumentVariable](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.API.Native.Document.CalculateDocumentVariable) event is triggered. A code within the event handler obtains the weather information. The [e.VariableName](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.CalculateDocumentVariableEventArgs.VariableName) property gets the name of the variable in the field, the <u>e.Arguments</u> property gets the location, and the <u>e.Value</u> property returns the calculated result.

The [MailMergeRecordStarted](https://docs.devexpress.com/OfficeFileAPI/DevExpress.XtraRichEdit.RichEditDocumentServer.MailMergeRecordStarted) event is handled to insert a hidden text that indicates when the document is created.

## Files to Review

* [Program.cs](.CS/WordProcessingFileAPI_CalcDocumentVariable/Program.cs) (VB: [Program.vb](./VB/WordProcessingFileAPI_CalcDocumentVariable/Program.vb))

## Documentation

* [DOCVARIABLE Field](https://docs.devexpress.com/OfficeFileAPI/15291/word-processing-document-api/fields/field-codes/docvariable)
* [Fields](https://docs.devexpress.com/OfficeFileAPI/15280/word-processing-document-api/fields)
* [How to: Replace a Placeholder with a Document Element](https://docs.devexpress.com/OfficeFileAPI/404369/word-processing-document-api/examples/search-and-replace/how-to-replace-a-placeholder-with-a-document-element)
