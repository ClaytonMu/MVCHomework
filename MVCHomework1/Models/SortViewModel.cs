using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCHomework1.Models
{
    public class SortViewModel
    {
        public string sortColumn { get; set; } = "";
        public string sortOrder { get; set; } = "ASC";
        public string keyword { get; set; } = "";
        public string jobTitle { get; set; } = "";
        public int page { get; set; } = 1;
    }
}