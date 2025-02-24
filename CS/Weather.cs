using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Net;
using System.Web;
using DevExpress.Utils;
using DevExpress.Docs.Text;

namespace WordProcessingFileAPI_CalcDocumentVariable
{
    class Weather
    {
        public static Dictionary<string, Conditions> weatherDic = new Dictionary<string, Conditions> ()
        {
             { "Berlin", new Conditions {Condition="Partly Cloudy", TempC="12", Humidity="82%", Wind="W 20km/h" }},
             { "Marseille", new Conditions {Condition="Clear", TempC="14", Humidity="67%", Wind="N 4km/h" }},
             { "Buenos Aires", new Conditions {Condition="Clear", TempC="10.4", Humidity="53%", Wind="NE 3.5km/h" }},
             { "London", new Conditions {Condition="Overcast", TempC="11", Humidity="82%", Wind="S 9.3km/h" }},
             { "Tula", new Conditions {Condition="Mist", TempC="0", Humidity="93%", Wind="ESE 7km/h" }}
        };

        public static Conditions GetCurrentConditions(string location)
        {
            Conditions result;
            if (weatherDic.TryGetValue(location, out result)) {
                return result;
            }
            else {
                return null;
            }
        }
    }
    public class Conditions
    {
        public string Condition {get;set;}
        public string TempC {get;set;}
        public string Humidity  {get;set;}
        public string Wind { get; set; }
    }
}
