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

        public StringField GetStringField()
        {
            var output = new StringField();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }
        public IntField GetIntField()
        {
            var output = new IntField();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }
        public Int32Field GetInt32Field()
        {
            var output = new Int32Field();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }
        public Int64Field GetInt64Field()
        {
            var output = new Int64Field();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }
        public DateTimeField GetDateTimeField()
        {
            var output = new DateTimeField();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }
        public BoolField GetBoolField()
        {
            var output = new BoolField();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }
        public BooleanField GetBooleanField()
        {
            var output = new BooleanField();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }
        public GuidField GetGuidField()
        {
            var output = new GuidField();
            output.Attribute = Attribute;
            output.FieldType = FieldType;
            output.Mandatory = Mandatory;
            output.Name = Name;
            output.ValidationType = ValidationType;
            return output;
        }

    }

    public class StringField : Field
    {
        public StringField()
        {
            FieldType = "String";
        }
        
    }
    public class IntField : Field
    {
        public IntField()
        {
            FieldType = "Int";
        }
    }
    public class Int32Field : Field
    {
        public Int32Field()
        {
            FieldType = "Int32";
        }
    }
    public class Int64Field : Field
    {
        public Int64Field()
        {
            FieldType = "Int64";
        }
    }
    public class DateTimeField : Field
    {
        public DateTimeField()
        {
            FieldType = "DateTime";
        }

    }
    public class BoolField : Field
    {
        public BoolField()
        {
            FieldType = "bool";
        }
    }
    public class BooleanField : Field
    {
        public BooleanField()
        {
            FieldType = "Boolean";
        }
    }
    public class GuidField : Field
    {
        public GuidField()
        {
            FieldType = "Guid";
        }
    }

}