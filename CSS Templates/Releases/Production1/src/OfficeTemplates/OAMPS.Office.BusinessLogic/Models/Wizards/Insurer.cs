﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OAMPS.Office.BusinessLogic.Interfaces.Wizards;

namespace OAMPS.Office.BusinessLogic.Models.Wizards
{
    public class Insurer : IInsurer
    {
        public string Title { get; set; }
        public string Id { get; set; }
        public string Category { get; set; }
    }
}
