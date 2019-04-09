Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Xml
Imports System.Net
Imports System.Web
Imports DevExpress.Utils
Imports DevExpress.Docs.Text

Namespace WordProcessingFileAPI_CalcDocumentVariable
	Friend Class Weather
		Public Shared weatherDic As New Dictionary(Of String, Conditions)() From {
			{"Berlin", New Conditions With {.Condition="Partly Cloudy", .TempC="12", .Humidity="82%", .Wind="W 20km/h"}
			},
			{"Marseille", New Conditions With {.Condition="Clear", .TempC="14", .Humidity="67%", .Wind="N 4km/h"}
			},
			{"Buenos Aires", New Conditions With {.Condition="Clear", .TempC="10.4", .Humidity="53%", .Wind="NE 3.5km/h"}
			},
			{"London", New Conditions With {.Condition="Overcast", .TempC="11", .Humidity="82%", .Wind="S 9.3km/h"}
			},
			{"Tula", New Conditions With {.Condition="Mist", .TempC="0", .Humidity="93%", .Wind="ESE 7km/h"}
			}
		}

		Public Shared Function GetCurrentConditions(ByVal location As String) As Conditions
			Dim result As Conditions = Nothing
			If weatherDic.TryGetValue(location, result) Then
				Return result
			Else
				Return Nothing
			End If
		End Function
	End Class
	Public Class Conditions
		Public Property Condition() As String
		Public Property TempC() As String
		Public Property Humidity() As String
		Public Property Wind() As String
	End Class
End Namespace
