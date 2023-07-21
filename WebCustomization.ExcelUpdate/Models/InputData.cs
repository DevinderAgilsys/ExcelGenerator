using Newtonsoft.Json;

namespace WebCustomization.ExcelUpdate
{
    partial class InputData
    {
        [JsonProperty("$content-type")]
        public string ContentType { get; set; }

        [JsonProperty("$content")]
        public string Content { get; set; }
    }
}
