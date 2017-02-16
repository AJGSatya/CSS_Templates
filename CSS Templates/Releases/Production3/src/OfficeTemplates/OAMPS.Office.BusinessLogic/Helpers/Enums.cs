using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OAMPS.Office.BusinessLogic.Helpers
{
   public static class Enums
    {
       public enum Segment
       {
           One,
           Two,
           Three,
           Four,
           Five,
           PersonalLines
       }

       public enum Remuneration
       {
           Fee,
           Commission,
           Combined
       }

       public enum Statutory
       {
           Wholesale,
           Retail,
           WholesaleWithRetail
       }

    }
}
