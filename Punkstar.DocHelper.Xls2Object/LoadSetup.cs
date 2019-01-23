using Newtonsoft.Json;
using System.Collections.Generic;

namespace Punkstar.DocHelper.Xls2Object
{

    [JsonObject]
    public class LoadSetup
    {
        [JsonProperty("Name")]
        public string Name { get; set; }
        [JsonProperty("Entities")]
        public List<Entity> Entities { get; set; }
        [JsonProperty("Regexs")]
        public List<RegexAttribute> Regexs { get; set; }
    }
}
