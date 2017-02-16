using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OAMPS.Office.Word.Helpers
{
   public static class Enums
    {
       public enum FormLoadType
       {
           OpenDocument,
           RibbonClick,
           RegenerateTemplate
       }

       public enum UsageTrackingType
       {
           NewDocument,
           UpdateData,
           RegenerateDocument 
       }
    }
}
