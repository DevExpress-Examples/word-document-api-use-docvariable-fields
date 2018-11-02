using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

namespace WordProcessingFileAPI_CalcDocumentVariable
{
    class SampleData : ArrayList
    {
        public SampleData()
        {
            Add(new AddresseeRecord("Maria", "Alfreds Futterkiste", "Obere Str. 57, Berlin", "Berlin"));
            Add(new AddresseeRecord("Laurence", "Bon app'", "12, rue des Bouchers, Marseille", "Marseille"));
            Add(new AddresseeRecord("Patricio", "Cactus Comidas para llevar", "Cerrito 333, Buenos Aires", "Buenos Aires"));
            Add(new AddresseeRecord("Thomas", "Around the Horn", "120 Hanover Sq., London", "London"));
            Add(new AddresseeRecord("Boris", "Express Developers", "Krasnoarmeiskiy prospect 25, Tula", "Tula"));
        }
    }

    public class AddresseeRecord
    {
        public string Name { get; set; }
        public string Company { get; set; }
        public string Address { get; set; }
        public string City { get; set; }

        public AddresseeRecord(string _Name, string _Company, string _Address, string _City)
        {
            this.Name = _Name;
            this.Company = _Company;
            this.Address = _Address;
            this.City = _City;
        }
    }
}
