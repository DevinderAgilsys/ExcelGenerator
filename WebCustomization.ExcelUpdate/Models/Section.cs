using System.Collections.Generic;

namespace WebCustomization.ExcelUpdate
{
    public class Section
    {
        public string SectionName { get; set; }
        public string SectionId { get; set; }
        public List<CardDetail> CardDetails { get; set; }
    }
}
