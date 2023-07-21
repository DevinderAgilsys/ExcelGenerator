using System.Collections.Generic;

namespace WebCustomization.ExcelUpdate
{
    public class CardDetail
    {
        public string CardName { get; set; }
        public string CardId { get; set; }
        public List<Field> Fields { get; set; }
    }
}
