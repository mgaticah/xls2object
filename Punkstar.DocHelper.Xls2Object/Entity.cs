using Newtonsoft.Json;
using System.Collections.Generic;

namespace Punkstar.DocHelper.Xls2Object
{
    [JsonObject]
    public class Entity
    {
        public Entity()
        {
            Entities = new List<Entity>();
            ConditionalEntities= new List<ConditionalEntity>();
            ConditionalFields = new List<ConditionalFields>();
            Fields = new List<Field>();
        }
        [JsonProperty("WorksheetName")]
        public string WorksheetName { get; set; }
        [JsonProperty("Name")]
        public string Name { get; set; }
        [JsonProperty("ClassName")]
        public string ClassName { get; set; }
        [JsonProperty("Parent")]
        public string Parent { get; set; }
        [JsonProperty("ExcelLookUpField")]
        public string ExcelLookUpField { get; set; }
        [JsonProperty("ParentLookUpField")]
        public string ParentLookUpField { get; set; }
        [JsonProperty("ParentAttribute")]
        public string ParentAttribute { get; set; }
        [JsonProperty("Fields")]
        public IList<Field> Fields { get; set; }
        [JsonProperty("ConditionalFields")]
        public IList<ConditionalFields> ConditionalFields { get; set; }
        [JsonProperty("ConditionalEntities")]
        public IList<ConditionalEntity> ConditionalEntities { get; set; }
        [JsonProperty("Entities")]
        public List<Entity> Entities { get; set; }
        [JsonProperty("ValidationType")]
        public string ValidationType;
        [JsonProperty("IsList")]
        public bool IsList;



    }
}