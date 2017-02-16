using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OAMPS.Office.Word.Helpers.Controls
{
   public class ValueRadioButton : RadioButton
    {
       //public string Value { get; set; }

       public Dictionary<string, string> Values { get; set; }

       public ValueRadioButton() 
       {
           Values = new Dictionary<string, string>();
       }

       public string GetValue(string key)
       {
           string returnValue = null;

           Values.TryGetValue(key, out returnValue);

           return returnValue;
       }

       public void AddValue(string key, string newValue)
       {
           string old = null;
           if (Values.TryGetValue(key, out old))
           {
               Values[key] = newValue;
           }
           else
           {
               Values.Add(key, newValue);
           }
       }

    }
}
