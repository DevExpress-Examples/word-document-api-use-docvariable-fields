<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/155868028/18.2.2%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830501)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Word Processing Document API - Use Document Variable (DOCVARIABLE) Fields

This example illustrates the use of a <strong>DOCVARIABLE</strong> field to provide additional information which is dependent on the value of a merged field. This technique is implemented so each merged document contains a weather report for a location that corresponds to the current data record.</p>
<p>NB: We do not provide code for retrieving weather information. You can implement a custom weather information provider.</p>
<p>The location is represented by a merge field. It is included as an argument within the DOCVARIABLE field. When the DOCVARIABLE field is updated, the <strong>DevExpress.XtraRichEdit.API.Native.Document.CalculateDocumentVariable</strong> event is triggered. A code within the event handler obtains the information on weather. It uses <u>e.VariableName</u> to get the name of the variable within the field, <u>e.Arguments</u> to get the location and returns the calculated result in <u>e.Value</u> property.<br /> The <strong>MailMergeRecordStarted</strong> event is handled to insert a hidden text indicating when the document is created. <br /> The <strong>MyProgressIndicatorService</strong> class is implemented and registered as a service to allow progress indication using the ProgressBar control.</p>
