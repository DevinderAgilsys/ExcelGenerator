using System.Collections.Generic;

namespace WebCustomization.ExcelUpdate
{
    public class Root
    {
        public string SheetName { get; set; }
        public string SheetId { get; set; }
        public bool IsCardView { get; set; }
        public List<CardDetail> CardDetails { get; set; }
        public List<Section> Sections { get; set; }
    }
}
