Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Collections

Namespace WordProcessingFileAPI_CalcDocumentVariable
	Friend Class SampleData
		Inherits ArrayList

		Public Sub New()
			Add(New AddresseeRecord("Maria", "Alfreds Futterkiste", "Obere Str. 57, Berlin", "Berlin"))
			Add(New AddresseeRecord("Laurence", "Bon app'", "12, rue des Bouchers, Marseille", "Marseille"))
			Add(New AddresseeRecord("Patricio", "Cactus Comidas para llevar", "Cerrito 333, Buenos Aires", "Buenos Aires"))
			Add(New AddresseeRecord("Thomas", "Around the Horn", "120 Hanover Sq., London", "London"))
			Add(New AddresseeRecord("Boris", "Express Developers", "Krasnoarmeiskiy prospect 25, Tula", "Tula"))
		End Sub
	End Class

	Public Class AddresseeRecord
		Public Property Name() As String
		Public Property Company() As String
		Public Property Address() As String
		Public Property City() As String

		Public Sub New(ByVal _Name As String, ByVal _Company As String, ByVal _Address As String, ByVal _City As String)
			Me.Name = _Name
			Me.Company = _Company
			Me.Address = _Address
			Me.City = _City
		End Sub
	End Class
End Namespace
