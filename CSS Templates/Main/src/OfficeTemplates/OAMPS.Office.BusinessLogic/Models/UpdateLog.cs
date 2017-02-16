using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OAMPS.Office.BusinessLogic.Models
{
    public class UpdateLog
    {
        public class UpdateLogEntry
        {
            public string Source { get; set; }
            public string Destination { get; set; }
            public DateTime LastModified { get; set; }
            public int Version { get; set; }
        }

        public List<UpdateLogEntry> Updates = new List<UpdateLogEntry>();
    }
}
