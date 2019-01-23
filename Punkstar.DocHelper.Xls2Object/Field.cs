using Newtonsoft.Json;

namespace Punkstar.DocHelper.Xls2Object
{
    [JsonObject]
    public class Field
    {
        public string Name { get; set; }
        public string Mandatory { get; set; }
        public string ValidationType { get; set; }
        public string Attribute { get; set; }
        public string FieldType { get; set; }

    }
}