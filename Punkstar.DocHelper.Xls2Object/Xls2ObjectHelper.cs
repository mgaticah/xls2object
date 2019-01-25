using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Punkstar.DocHelper.Xls2Object
{
    public class Xls2ObjectHelper
    {
        private const int fieldNameRowNumber = 2;
        private const int maxExcelRowsExpected = 100000;
        private const string ExcelColumnLookupDefaultText = "Excel column lookup";
        private List<EntityRange> EntitiesRanges;
        public ImportSettings setup;
        public Assembly assembly;
        public void AssignStringValue(StringField field, object instance, PropertyInfo prop,  EntityRange range, string value, int row)
        {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna '{0}' con el valor para el atributo '{1}' (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                prop.SetValue(instance, "", null);
                return;
            }

            if (setup.Regexs != null)
            {
                var regex = setup.Regexs.FirstOrDefault(x => x.Attribute == field.Attribute && x.ClassName == range.ClassName) != null ? setup.Regexs.FirstOrDefault(x => x.Attribute == field.Attribute && x.ClassName == range.ClassName).Expression : "";
                if (regex != null && !string.IsNullOrEmpty(regex) && (Regex.Match(value, regex).Value == "" || string.IsNullOrEmpty(value)))
                {
                    throw new Exception(string.Format("ERROR(Linea {2}) Entidad '{5}': Columna '{0}' con el valor para el atributo '{1}' (tipo: '{3}') no cumple con la expresión regular '{4}'", field.Name, field.Attribute, row, field.FieldType, regex, range.Name));
                }
            }
            prop.SetValue(instance, value, null);
        }
        public void AssignIntValue(IntField field, object instance, PropertyInfo prop,  EntityRange range, string value, int row) {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad '{4}': Columna '{0}' con el valor para el atributo '{1}' (tipo: '{3}') es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                try
                {
                    prop.SetValue(instance, null, null);
                }
                catch (Exception)
                {
                    prop.SetValue(instance, 0, null);
                }
                return;
            }
            int n;
            bool isNumeric = int.TryParse(value, out n);
            if (!isNumeric)
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
            }

            prop.SetValue(instance, n, null);
        }
        public void AssignInt32Value(Int32Field field, object instance, PropertyInfo prop,  EntityRange range, string value, int row) {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad '{4}': Columna '{0}' con el valor para el atributo '{1}' (tipo: '{3}') es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                try
                {
                    prop.SetValue(instance, null, null);
                }
                catch (Exception)
                {
                    prop.SetValue(instance, 0, null);
                }
                return;
            }
            int n;
            bool isNumeric = int.TryParse(value, out n);
            if (!isNumeric)
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
            }

            prop.SetValue(instance, n, null);
        }
        public void AssignInt64Value(Int64Field field, object instance, PropertyInfo prop,  EntityRange range, string value, int row) {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad '{4}': Columna '{0}' con el valor para el atributo '{1}' (tipo: '{3}') es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                try
                {
                    prop.SetValue(instance, null, null);
                }
                catch (Exception)
                {
                    prop.SetValue(instance, 0, null);
                }
                return;
            }
            int n;
            bool isNumeric = int.TryParse(value, out n);
            if (!isNumeric)
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
            }

            prop.SetValue(instance, n, null);
        }
        public void AssignDateTimeValue(DateTimeField field, object instance, PropertyInfo prop, EntityRange range, string value, int row) {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                prop.SetValue(instance, new DateTime(), null);
                return;
            }

            DateTime date;
            if (!DateTime.TryParse(value, out date))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
            }

            prop.SetValue(instance, date, null);
        }
        public void AssignBoolValue(BoolField field, object instance, PropertyInfo prop,  EntityRange range, string value, int row) {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                prop.SetValue(instance, false, null);
                return;
            }
            bool varBool;
            if (!bool.TryParse(value, out varBool))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
            }

            prop.SetValue(instance, varBool, null);
        }
        public void AssignBooleanValue(BooleanField field, object instance, PropertyInfo prop,  EntityRange range, string value, int row) {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                prop.SetValue(instance, false, null);
                return;
            }
            bool varBool;
            if (!bool.TryParse(value, out varBool))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
            }

            prop.SetValue(instance, varBool, null);
        }
        public void AssignGuidValue(GuidField field, object instance, PropertyInfo prop,  EntityRange range, string value, int row) {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            Guid varGuid;
            if (!Guid.TryParse(value, out varGuid))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
            }

            prop.SetValue(instance, varGuid, null);
        }
        public void AssignAnotherKindOfValue(GuidField field, object instance, PropertyInfo prop, EntityRange range, string value, int row)
        {
            if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
            {
                throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
            }

            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            var childType = assembly.GetTypes().FirstOrDefault(x => x.Name == field.FieldType);
            var childInstance = Activator.CreateInstance(childType);
            var parseDirecto = true;
            if (childType.IsEnum)
            {
                try
                {
                    childInstance = Enum.Parse(childType, value);
                    prop.SetValue(instance, childInstance, null);
                }
                catch (Exception)
                {
                    parseDirecto = false;
                }
                if (!parseDirecto)
                {
                    try
                    {
                        var encontrado = false;
                        foreach (var declaredField in ((TypeInfo)childType).DeclaredFields)
                        {
                            foreach (var customAttribute in declaredField.CustomAttributes)
                            {
                                if (customAttribute.ConstructorArguments != null && customAttribute.ConstructorArguments.Count > 0)
                                {
                                    if (customAttribute.ConstructorArguments.Any(x => x.Value.ToString().ToLower() == value.ToLower()))
                                    {
                                        childInstance = Enum.Parse(childType, declaredField.Name);
                                        prop.SetValue(instance, childInstance, null);
                                        encontrado = true;
                                        break;
                                    }
                                }

                                if (encontrado)
                                {
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("No ha sido posible interretar el valor '{0}' en el enumerador '{1}.Error:{2}'", value, field.FieldType, ex.Message));
                    }
                }
            }
            else if (childType.IsClass)
            {
                var childInstanceProp = instance.GetType().GetProperties().FirstOrDefault(x => x.GetType() == childType);
                if (childInstanceProp != null)
                {
                    childInstanceProp.SetValue(childInstance, value, null);
                    prop.SetValue(instance, childInstance, null);
                }
            }
        }
        public void AssignValue(object instance, PropertyInfo prop, Field field, EntityRange range, string value, int row)
        {
            try
            {
                switch (field.FieldType.ToLower())
                {
                    case "string":
                        AssignStringValue(field.GetStringField(), instance, prop, range, value, row);
                        break;
                    case "int":
                        AssignIntValue(field.GetIntField(), instance, prop, range, value, row);
                        break;
                    case "int32":
                        AssignInt32Value(field.GetInt32Field(), instance, prop, range, value, row);
                        break;
                    case "int64":
                        AssignInt64Value(field.GetInt64Field(), instance, prop, range, value, row);
                        break;
                    case "datetime":
                        AssignDateTimeValue(field.GetDateTimeField(), instance, prop, range, value, row);
                        break;
                    case "bool":
                        AssignBoolValue(field.GetBoolField(), instance, prop, range, value, row);
                        break;
                    case "boolean":
                        AssignBooleanValue(field.GetBooleanField(), instance, prop, range, value, row);
                        break;
                    case "guid":
                        AssignGuidValue(field.GetGuidField(),instance, prop,  range, value, row);
                        break;
                    default:
                        AssignAnotherKindOfValue(field.GetGuidField(), instance, prop, range, value, row);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public ImportSettings CreateImportSettingsByObject(object instance, int deepLevel, string[] excludedClasses)
        {
            var output = new ImportSettings();
            var instanceType = instance.GetType();
            output.Name = instanceType.Name;
            output.Entities = new List<Entity>();
            var mainEntity = GetMainEntity(instance, excludedClasses);
            var dependantEntities = GetDependantEntities(instance, 0, deepLevel, excludedClasses);
            dependantEntities.Select(x => x.Parent = instanceType.Name);
            foreach (var entity in dependantEntities)
            {
                entity.Parent = instanceType.FullName;
            }

            mainEntity.Entities.AddRange(dependantEntities);
            output.Entities.Add(mainEntity);
            return output;
        }
        public string CreateImportSettingsJsonByClass(object instance, int deepLevel, string[] excludedClasses)
        {
            var output = CreateImportSettingsByObject(instance, deepLevel, excludedClasses);
            return JsonConvert.SerializeObject(output);
        }
        public bool Evaluate(Condition condition, object instance)
        {
            if (string.IsNullOrEmpty(condition.Entity))
            {
                var properties = instance.GetType().GetProperties();
                var instanceProperty = instance.GetType().GetProperty(condition.Field);
                if (condition.Operation == Enums.Operator.Equal)
                {
                    return instanceProperty.GetValue(instance, null).ToString() == condition.Value;
                }

                if (condition.Operation == Enums.Operator.GreaterThan)
                {
                    return double.Parse(instanceProperty.GetValue(instance, null).ToString()) > double.Parse(condition.Value);
                }

                if (condition.Operation == Enums.Operator.LesserThan)
                {
                    return double.Parse(instanceProperty.GetValue(instance, null).ToString()) < double.Parse(condition.Value);
                }
            }
            return false;
        }
        public string GetCellValue(ExcelWorksheet workSheet, EntityRange range, Field field, int row)
        {
            for (var columnNumber = range.Start; columnNumber <= range.End; columnNumber++)
            {
                if (workSheet.Cells[fieldNameRowNumber, columnNumber].Text.ToLower() == field.Name.ToLower())
                {
                    return workSheet.Cells[row, columnNumber].Value.ToString();
                }
            }

            return "";
        }
        public string GetCellValue(ExcelWorksheet workSheet, EntityRange range, string columnName, int row)
        {
            for (var columnNumber = range.Start; columnNumber <= range.End; columnNumber++)
            {
                if (workSheet.Cells[fieldNameRowNumber, columnNumber].Text.ToLower() == columnName.ToLower())
                {
                    return workSheet.Cells[row, columnNumber].Text;
                }
            }

            return "";
        }
        public List<Entity> GetDependantEntities(object instance, int deep, int maxDeep, string[] excludedClasses)
        {
            var entities = new List<Entity>();
            var instanceType = instance.GetType();
            var fields = new List<Field>();
            foreach (var instanceProperty in instanceType.GetProperties())
            {
                if (!instanceProperty.CanWrite || !instanceProperty.GetSetMethod(true).IsPublic)
                {
                    continue;
                }

                var excluded = false;
                if (excludedClasses != null)
                {
                    foreach (var excludedClass in excludedClasses)
                    {
                        if (instanceProperty.PropertyType.FullName.Contains(excludedClass))
                        {
                            excluded = true;
                            break;
                        }
                    }
                }

                if (excluded)
                {
                    continue;
                }

                string propertyTypeName = instanceProperty.PropertyType.Name.ToLower();
                if (propertyTypeName.Contains("nullable"))
                {
                    propertyTypeName = instanceProperty.PropertyType.GenericTypeArguments[0].Name.ToLower();
                }

                switch (propertyTypeName)
                {
                    case "string":
                    case "int":
                    case "int16":
                    case "int32":
                    case "int64":
                    case "datetime":
                    case "bool":
                    case "boolean":
                    case "guid":
                    case "byte":
                        break;
                    default:

                        if (instanceProperty.PropertyType.Name.Contains("List"))
                        {
                            object listInstance = GetListInstance(instanceProperty);
                            List<Entity> listDependantEntities = GetListInstanceDependantEntities(deep, maxDeep, excludedClasses, listInstance);
                            Entity entity = GetListEntity(excludedClasses, instanceProperty, listInstance, listDependantEntities);
                            entities.Add(entity);
                        }
                        else
                        {
                            var propertyInstance = Activator.CreateInstance(instanceProperty.PropertyType);
                            Entity propertyEntity = GetPropertyEntity(excludedClasses, instanceType, instanceProperty, propertyInstance);
                            var propertyDependantEntities = new List<Entity>();
                            if (deep + 1 < maxDeep)
                            {
                                propertyDependantEntities = GetDependantEntities(propertyInstance, deep + 1, maxDeep, excludedClasses);
                                foreach (var propertyDependantEntity in propertyDependantEntities)
                                {
                                    propertyDependantEntity.Parent = instanceProperty.Name;

                                }
                            }
                            propertyEntity.Entities = propertyDependantEntities;
                            entities.Add(propertyEntity);
                        }
                        break;
                }
            }
            return entities;
        }
        private Entity GetPropertyEntity(string[] excludedClasses, Type instanceType, PropertyInfo instanceProperty, object propertyInstance)
        {
            var propertyEntity = GetMainEntity(propertyInstance, excludedClasses);
            propertyEntity.Name = instanceProperty.Name;
            propertyEntity.ParentAttribute = instanceProperty.Name;
            propertyEntity.ExcelLookUpField = string.Format("'{0}' Excel column lookup ", instanceProperty.Name);
            propertyEntity.ParentLookUpField = "field in parent to look up for";
            propertyEntity.IsList = false;
            propertyEntity.Parent = instanceType.FullName;
            return propertyEntity;
        }
        public List<Field> GetEntityFields(object instance, string[] excludedClasses)
        {
            var instanceType = instance.GetType();
            if (instanceType.Name.Contains("List"))
            {
                instanceType = instance.GetType().GenericTypeArguments[0];
            }

            var fields = new List<Field>();
            foreach (var property in instanceType.GetProperties())
            {
                Field field = null;
                switch (property.PropertyType.Name.ToLower())
                {
                    case "string":
                        field = new StringField();
                        break;
                    case "int":
                        field = new IntField();
                        break;
                    case "int32":
                        field = new Int32Field();
                        break;
                    case "int64":
                        field = new Int64Field();
                        break;
                    case "datetime":
                        field = new DateTimeField();
                        break;
                    case "bool":
                        field = new BoolField();
                        break;
                    case "boolean":
                        field = new BooleanField();
                        break;
                    case "guid":
                        field = new GuidField();
                        break;
                    default:
                        field = new Field();
                        break;
                }
                MethodInfo setMethod = property.GetSetMethod();
                if (setMethod != null)
                {
                    var excluded = false;
                    if (excludedClasses != null)
                    {
                        foreach (var excludedClass in excludedClasses)
                        {
                            if (property.PropertyType.FullName.Contains(excludedClass))
                            {
                                excluded = true;
                                break;
                            }
                        }
                    }

                    if (excluded)
                    {
                        continue;
                    }
                    field.Attribute = property.Name;
                    field.Mandatory = "false";
                    field.Name = "Excel column name";
                    field.ValidationType = "";
                    fields.Add(field);
                }
            }
            return fields;
        }
        public List<EntityRange> GetEntityRanges(List<Entity> Entities, ExcelPackage XLSFile)
        {
            try
            {
                var currentSheet = XLSFile.Workbook.Worksheets;
                List<EntityRange> validatedEntityRanges = GetValidatedEntityRanges(Entities, currentSheet);
                List<EntityRange> subEntityRanges = GetSubEntitiesRanges(Entities, XLSFile, validatedEntityRanges);
                if (subEntityRanges.Count > 0)
                {
                    validatedEntityRanges.AddRange(subEntityRanges);
                }

                foreach (var entity in Entities.FirstOrDefault().Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList())
                {
                    if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.WorksheetName)))
                    {
                        validatedEntityRanges.AddRange(GetEntityRanges(entity.Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList(), XLSFile));
                    }
                }
                return validatedEntityRanges;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        public ImportSettings GetImportSettings(Stream JSONFile)
        {
            ImportSettings importSettings = null;
            using (var reader = new StreamReader(JSONFile))
            {
                StreamReader readStream = new StreamReader(JSONFile, Encoding.UTF8);
                string jsonString = "";
                jsonString = jsonString + readStream.ReadToEnd();
                jsonString = jsonString.Replace(@"\", " ");
                importSettings = JsonConvert.DeserializeObject<ImportSettings>(jsonString);
            }
            return importSettings;
        }
        private static object GetListInstance(PropertyInfo instanceProperty)
        {
            var listType = typeof(List<>);
            var constructedListType = listType.MakeGenericType(instanceProperty.PropertyType.GenericTypeArguments[0]);
            var listInstance = Activator.CreateInstance(constructedListType);
            return listInstance;
        }
        private Entity GetListEntity(string[] excludedClasses, PropertyInfo instanceProperty, object listInstance, List<Entity> listDependantEntities)
        {
            var entity = new Entity();
            entity = GetMainEntity(listInstance, excludedClasses);
            entity.Name = instanceProperty.Name;
            entity.ParentAttribute = instanceProperty.Name;
            entity.IsList = true;
            entity.ExcelLookUpField = string.Format("'{0}' Excel column lookup ", instanceProperty.Name);
            entity.ParentLookUpField = "field in parent to look up for";
            entity.Entities = listDependantEntities;
            return entity;
        }
        private List<Entity> GetListInstanceDependantEntities(int deep, int maxDeep, string[] excludedClasses, object listInstance)
        {
            var dependantEntities = new List<Entity>();
            if (listInstance.GetType().GenericTypeArguments.Count() > 0)
            {
                var testObject = Activator.CreateInstance(listInstance.GetType().GenericTypeArguments[0]);
                if (deep < maxDeep)
                {
                    dependantEntities = GetDependantEntities(testObject, deep + 1, maxDeep, excludedClasses);
                }
            }

            return dependantEntities;
        }
        private List<EntityRange> GetSubEntitiesRanges(List<Entity> Entities, ExcelPackage XLSFile, List<EntityRange> validatedEntityRanges)
        {
            var subEntityRanges = new List<EntityRange>();
            foreach (var entityRange in validatedEntityRanges)
            {
                var entity = Entities.FirstOrDefault(x => x.Name == entityRange.Name);
                if (entity != null)
                {
                    var subEntities = entity.Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList();
                    if (subEntities != null && subEntities.Count > 0)
                    {
                        subEntityRanges.AddRange(GetEntityRanges(subEntities, XLSFile));
                    }
                }
            }
            return subEntityRanges;
        }
        private static List<EntityRange> GetValidatedEntityRanges(List<Entity> Entities, ExcelWorksheets currentSheet)
        {
            var entityRanges = new List<EntityRange>();
            foreach (var entity in Entities)
            {
                if (entity.WorksheetName == null && entity.Parent == null)
                {
                    continue;
                }

                var worksheet = currentSheet.FirstOrDefault(x => x.Name == entity.WorksheetName);
                if (worksheet == null)
                {
                    throw new Exception(string.Format("Worksheet '{0}' not found.", entity.WorksheetName));
                }

                var numberOfColumns = worksheet.Dimension.End.Column;
                var lastFoundEntity = "";
                for (int columnNumber = 1; columnNumber <= numberOfColumns; columnNumber++)
                {
                    if (worksheet.Cells[1, columnNumber].Text != "")
                    {
                        lastFoundEntity = worksheet.Cells[1, columnNumber].Text;
                        if (lastFoundEntity == "")
                        {
                            entityRanges.Add(new EntityRange { SpreadSheetName = entity.WorksheetName, Name = lastFoundEntity, Start = columnNumber, ClassName = Entities.Any(x => x.Name == lastFoundEntity) ? Entities.FirstOrDefault(x => x.Name == lastFoundEntity).ClassName : "NoName" });
                        }
                        else if (lastFoundEntity != worksheet.Cells[1, columnNumber].Text)
                        {
                            entityRanges.Last().End = columnNumber - 1;
                        }
                    }
                }
                if (entityRanges.Count == 0)
                {
                    entityRanges.Add(new EntityRange { SpreadSheetName = entity.WorksheetName, Name = lastFoundEntity, Start = 1, ClassName = Entities.Any(x => x.Name == lastFoundEntity) ? Entities.FirstOrDefault(x => x.Name == lastFoundEntity).ClassName : "NoName" });
                }
                entityRanges.Last().End = numberOfColumns;
            }

            return entityRanges;
        }
        public Entity GetMainEntity(object instance, string[] excludedClasses)
        {
            var instanceType = instance.GetType();
            var mainEntity = new Entity();
            if (instance.GetType().FullName.Contains("List"))
            {
                mainEntity.ClassName = instanceType.GenericTypeArguments[0].FullName;
                mainEntity.Name = instanceType.GenericTypeArguments[0].FullName;
            }
            else
            {
                mainEntity.ClassName = instanceType.FullName;
                mainEntity.Name = instanceType.Name;
            }
            mainEntity.Fields = GetEntityFields(instance, excludedClasses);
            return mainEntity;
        }
        public List<dynamic> GetObjectsFromExcel(string excelBase64String, Stream JSONFile)
        {
            var data = Convert.FromBase64String(excelBase64String);
            var stream = new MemoryStream(data);
            return GetObjectsFromExcel(stream, JSONFile);
        }
        public List<dynamic> GetObjectsFromExcel(Stream excelStream, Stream JSONFile)
        {
            return GetObjectsFromExcel(new ExcelPackage(excelStream), JSONFile);
        }
        public List<dynamic> GetObjectsFromExcel(dynamic parentObject, ExcelPackage XLSFile, List<Entity> Entities, List<EntityRange> EntityRanges)
        {
            var list = new List<dynamic>();
            var currentSheet = XLSFile.Workbook.Worksheets;
            foreach (var entity in Entities)
            {
                if (entity.WorksheetName == null && entity.Parent == null)
                {
                    var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                    var parentInstance = Activator.CreateInstance(type);
                    if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.WorksheetName)))
                    {
                        list.AddRange(GetObjectsFromExcel(parentInstance, XLSFile, entity.Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList(), EntityRanges));
                    }

                    continue;
                }
                var workSheet = currentSheet.FirstOrDefault(x => x.Name == entity.WorksheetName);
                int worksheetRowsNumber = WorksheetRowsNumber(workSheet);
                for (int currentRowNumber = 3; currentRowNumber <= worksheetRowsNumber; currentRowNumber++)
                {
                    var range = EntitiesRanges.FirstOrDefault(x => x.Name == entity.Name);
                    if (range == null)
                    {
                        throw new Exception(string.Format("Was not found range for entity '{0}' at worksheet '{1}' at section '{2}'", entity.ClassName, entity.WorksheetName, entity.Name));
                    }

                    var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                    var instance = Activator.CreateInstance(type);
                    PopulateInstance(instance, Entities.FirstOrDefault(x => x.Name == entity.Name), workSheet, range, currentRowNumber);
                    if (Entities.FirstOrDefault(x => x.Name == entity.Name).Entities.Any(x => !string.IsNullOrEmpty(x.WorksheetName)))
                    {
                        instance = GetObjectsFromExcel(instance, XLSFile, Entities.FirstOrDefault(x => x.Name == entity.Name).Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList(), EntitiesRanges).FirstOrDefault();
                    }

                    PropertyInfo[] parentPropertiesList = parentObject.GetType().GetProperties();
                    var excelLookupValue = GetCellValue(workSheet, range, entity.ExcelLookUpField, currentRowNumber);
                    var parentType = parentObject.GetType();
                    var parentLookupProperty = parentType.GetProperty(entity.ParentLookUpField);
                    if (entity.ExcelLookUpField != "" && !entity.ExcelLookUpField.Contains(ExcelColumnLookupDefaultText) && string.IsNullOrWhiteSpace(excelLookupValue))
                    {
                        throw new Exception(string.Format("Cell '{0}' has no value or doesnt exist at worksheet {1}", entity.ExcelLookUpField, entity.WorksheetName));
                    }

                    var parentAttributePropertyInfo = parentPropertiesList.FirstOrDefault(x => x.Name.Equals(entity.ParentAttribute));
                    if (parentAttributePropertyInfo == null)
                    {
                        throw new Exception(string.Format("'{0}' is not an attribute of '{1}' class", entity.ParentAttribute, entity.Parent));
                    }

                    SetParentAttribute(parentObject, entity, instance, excelLookupValue, parentLookupProperty, parentAttributePropertyInfo);
                }
            }
            list.Add(parentObject);
            return list;

        }

        public List<dynamic> GetObjectsFromExcel(ExcelPackage XLSFile, Stream JSONFile)
        {
            var list = new List<dynamic>();
            setup = GetImportSettings(JSONFile);
            var entities = new List<Entity>();
            entities = setup.Entities;
            EntitiesRanges = GetEntityRanges(entities, XLSFile);
            foreach (var entity in entities)
            {
                var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                var parentInstance = Activator.CreateInstance(type);

                if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.WorksheetName)))
                {
                    list.AddRange(GetObjectsFromExcel(parentInstance, XLSFile, entity.Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList(), EntitiesRanges));
                }
                else
                {
                    list.AddRange(GetObjectsFromExcel(entity, XLSFile, EntitiesRanges));
                }
            }
            return list;
        }

        private IEnumerable<dynamic> GetObjectsFromExcel(Entity entity, ExcelPackage XLSFile, List<EntityRange> entitiesRanges)
        {
            var list = new List<dynamic>();
            var currentSheet = XLSFile.Workbook.Worksheets;

            if (entity.WorksheetName == null && entity.Parent == null)
            {
                var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                var parentInstance = Activator.CreateInstance(type);
                if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.WorksheetName)))
                {
                    list.AddRange(GetObjectsFromExcel(parentInstance, XLSFile, entity.Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList(), entitiesRanges));
                }
            }
            var workSheet = currentSheet.FirstOrDefault(x => x.Name == entity.WorksheetName);
            int worksheetRowsNumber = WorksheetRowsNumber(workSheet);
            for (int currentRowNumber = 3; currentRowNumber <= worksheetRowsNumber; currentRowNumber++)
            {
                var range = EntitiesRanges.FirstOrDefault(x => x.Name == entity.Name);
                if (range == null)
                {
                    throw new Exception(string.Format("Was not found range for entity '{0}' at worksheet '{1}' at section '{2}'", entity.ClassName, entity.WorksheetName, entity.Name));
                }
                var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                var instance = Activator.CreateInstance(type);
                PopulateInstance(instance, entity, workSheet, range, currentRowNumber);
                if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.WorksheetName)))
                {
                    instance = GetObjectsFromExcel(instance, XLSFile, entity.Entities.Where(x => !string.IsNullOrEmpty(x.WorksheetName)).ToList(), EntitiesRanges).FirstOrDefault();
                }
                list.Add(instance);
            }
            return list;
        }

        private void SetParentAttribute(dynamic parentObject, Entity entity, object instance, string excelLookupValue, dynamic parentLookupProperty, PropertyInfo parentAttributePropertyInfo)
        {
            if (parentAttributePropertyInfo.PropertyType.IsGenericType && parentAttributePropertyInfo.PropertyType.GetGenericTypeDefinition().Name.Contains("List"))
            {
                SetParentListAttribute(parentObject, entity, instance, excelLookupValue, parentLookupProperty, parentAttributePropertyInfo);

            }
            else if (parentLookupProperty != null && parentLookupProperty.Name != "")
            {
                if (parentObject.GetType().GetProperty(parentLookupProperty.Name).GetValue(parentObject, null) == excelLookupValue)
                {
                    parentAttributePropertyInfo.SetValue(parentObject, instance, null);
                }
            }
            else
            {
                parentAttributePropertyInfo.SetValue(parentObject, instance, null);
            }
        }
        public void LoadAssembly(Stream inputStream)
        {
            if (inputStream == null)
            {
                return;
            }

            byte[] data;

            var memoryStream = inputStream as MemoryStream;
            if (memoryStream == null)
            {
                memoryStream = new MemoryStream();
                inputStream.CopyTo(memoryStream);
            }
            data = memoryStream.ToArray();
            assembly = AppDomain.CurrentDomain.Load(data);
        }
        public void LoadAssembly(string assemblyName, Stream inputStream)
        {
            if (!AppDomain.CurrentDomain.GetAssemblies().Any(x => x.ManifestModule.ScopeName == assemblyName))
            {
                LoadAssembly(inputStream);
            }
            else
            {
                assembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(x => x.ManifestModule.ScopeName == assemblyName);
            }
        }
        public object PopulateAttribute(object instance, ExcelWorksheet workSheet, EntityRange range, Field field, int row)
        {
            var attributePropertyInfo = instance.GetType().GetProperties().FirstOrDefault(x => x.Name.Equals(field.Attribute));
            if (attributePropertyInfo == null)
            {
                throw new Exception(string.Format("El atributo {0} no encontrado en la clase {1}. Verifique el assembly", field.Attribute, instance.GetType()));
            }

            if (null != attributePropertyInfo && attributePropertyInfo.CanWrite)
            {
                var value = GetCellValue(workSheet, range, field, row);
                AssignValue(instance, attributePropertyInfo, field, range, value, row);
                ValidateField(instance, field, value, row);
            }
            return instance;
        }
        public object PopulateInstance(object instance, Entity entity, ExcelWorksheet workSheet, EntityRange range, int row)
        {
            foreach (var field in entity.Fields)
            {
                instance = PopulateAttribute(instance, workSheet, range, field, row);
            }
            if (entity.ConditionalEntities != null)
            {
                foreach (var _ConditionalEntity in entity.ConditionalEntities)
                {
                    var IsRequired = true;
                    foreach (var condition in _ConditionalEntity.Conditions)
                    {
                        if (!Evaluate(condition, instance))
                        {
                            IsRequired = false;
                            break;
                        }
                    }
                    if (IsRequired)
                    {
                        foreach (var _Entity in _ConditionalEntity.Entities)
                        {
                            var conditionalRange = EntitiesRanges.FirstOrDefault(x => x.Name == _Entity.Name);
                            var conditionalType = assembly.GetTypes().FirstOrDefault(x => x.FullName == _Entity.ClassName);
                            var conditionalInstance = Activator.CreateInstance(conditionalType);
                            var childInstance = PopulateInstance(conditionalInstance, _Entity, workSheet, conditionalRange, row);
                            var attributePropertyInfo = instance.GetType().GetProperty(_ConditionalEntity.Attribute, BindingFlags.Public | BindingFlags.Instance);
                            if (attributePropertyInfo == null)
                            {
                                throw new Exception(string.Format("Attribute '{0}' not found in class: '{1}'. Verify Assembly", _ConditionalEntity.Attribute, conditionalInstance.GetType()));
                            }

                            attributePropertyInfo.SetValue(instance, childInstance, null);
                        }
                    }
                }
            }

            ValidateEntity(entity, instance, row);
            return instance;
        }
        public object PopulateObjectWithRandomData(object objectInstance)
        {
            if (objectInstance == null)
            {
                return null;
            }

            foreach (var field in objectInstance.GetType().GetProperties())
            {
                if (!field.CanWrite || !field.GetSetMethod(true).IsPublic)
                {
                    continue;
                }

                string propertyTypeName = field.PropertyType.Name.ToLower();
                if (propertyTypeName.Contains("nullable"))
                {
                    propertyTypeName = field.PropertyType.GenericTypeArguments[0].Name.ToLower();
                }

                switch (propertyTypeName)
                {
                    case "string":
                        field.SetValue(objectInstance, RandomString(20), null);
                        break;
                    case "int":
                        field.SetValue(objectInstance, random.Next(1, 10000), null);
                        break;
                    case "int16":
                        field.SetValue(objectInstance, (Int16)random.Next(1, 100), null);
                        break;
                    case "int32":
                        field.SetValue(objectInstance, random.Next(1, 10000), null);
                        break;
                    case "int64":
                        field.SetValue(objectInstance, (Int64)random.Next(1, 10000), null);
                        break;
                    case "datetime":
                        field.SetValue(objectInstance, DateTime.Now, null);
                        break;
                    case "bool":
                        field.SetValue(objectInstance, true, null);
                        break;
                    case "boolean":
                        field.SetValue(objectInstance, true, null);
                        break;
                    case "guid":
                        field.SetValue(objectInstance, Guid.NewGuid(), null);
                        break;
                    case "byte":
                        field.SetValue(objectInstance, new byte(), null);
                        break;
                    default:
                        if (field.PropertyType.Name.Contains("List") || typeof(IEnumerable).IsAssignableFrom(objectInstance.GetType()))
                        {
                            var listType = typeof(List<>);
                            if (field.PropertyType.GenericTypeArguments.Count() > 0)
                            {
                                var constructedListType = listType.MakeGenericType(field.PropertyType.GenericTypeArguments[0]);
                                var listInstance = PopulateObjectWithRandomData(Activator.CreateInstance(constructedListType));
                                field.SetValue(objectInstance, listInstance);
                            }
                            else
                            {
                                try
                                {
                                    Type objTyp = objectInstance.GetType();
                                    var IListRef = typeof(List<>);
                                    Type[] IListParam = { objTyp };
                                    object Result = Activator.CreateInstance(IListRef.MakeGenericType(IListParam));
                                    var objTemp = Activator.CreateInstance(objTyp);

                                    Result.GetType().GetMethod("Add").Invoke(Result, new[] { objTemp });

                                }
                                catch (Exception)
                                {
                                    //TODO: control this
                                }
                            }
                        }
                        else
                        {
                            var childType = assembly.GetTypes().FirstOrDefault(x => x.Name == field.PropertyType.Name);
                            var childobjectInstance = Activator.CreateInstance(childType);
                            if (childobjectInstance.GetType().IsEnum)
                            {
                                field.SetValue(objectInstance, childobjectInstance.GetType().GetEnumValues().GetValue(0), null);
                            }
                            else
                            {
                                field.SetValue(objectInstance, childobjectInstance, null);
                            }
                        }
                        break;
                }
            }
            return objectInstance;
        }
        private static Random random = new Random();
        private static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
        public void SetCellValue(string excelFilePath, string newExcelFilePath, Dictionary<string, Dictionary<string, string>> ValuesToSet)
        {
            var package = new ExcelPackage(new FileInfo(excelFilePath));
            foreach (var spreadsheet in ValuesToSet)
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets[spreadsheet.Key];
                foreach (var cell in spreadsheet.Value)
                {
                    workSheet.Cells[cell.Key].Value = cell.Value;
                }
            }
            package.SaveAs(new FileInfo(newExcelFilePath));
            package.Dispose();
        }
        private static void SetParentListAttribute(dynamic parentObject, Entity entity, object instance, string excelLookupValue, dynamic parentLookupProperty, PropertyInfo parentAttributePropertyInfo)
        {
            if (parentLookupProperty != null && parentLookupProperty.Name != "")
            {
                if (parentObject.GetType().GetProperty(parentLookupProperty.Name).GetValue(parentObject, null) == excelLookupValue)
                {
                    IList reflectedList = (IList)parentAttributePropertyInfo.GetValue(parentObject);
                    reflectedList.Add(instance);
                    parentObject.GetType().GetProperty(entity.ParentAttribute).SetValue(parentObject, reflectedList);
                }
            }
            else
            {
                IList reflectedList = (IList)parentAttributePropertyInfo.GetValue(parentObject);
                reflectedList.Add(instance);
                parentObject.GetType().GetProperty(entity.ParentAttribute).SetValue(parentObject, reflectedList);
            }
        }
        public void ValidateEntity(Entity entity, object instance, int row)
        {
            if (!string.IsNullOrEmpty(entity.ValidationType))
            {
                MethodInfo method = instance.GetType().GetMethod(entity.ValidationType);
                if (method == null)
                {
                    throw new Exception(string.Format("Validation method '{0}' not found in class:'{1}'. Verify assembly", entity.ValidationType, entity.ClassName));
                }

                object result = method.Invoke(instance, null);
                if (!((bool)result))
                {
                    throw new Exception(string.Format("ERROR Row '{0}' doesnt satisfy validation rule '{1}'", row, entity.ValidationType));
                }
            }
        }
        public void ValidateField(object instance, Field field, string value, int row)
        {
            if (!string.IsNullOrEmpty(field.ValidationType))
            {
                MethodInfo method = instance.GetType().GetMethod(field.ValidationType);
                if (method == null)
                {
                    throw new Exception(string.Format("Validation method '{0}' was not found in class '{1}'. Verify assembly", field.ValidationType, instance.GetType()));
                }

                object result = method.Invoke(instance, null);
                if (!((bool)result))
                {
                    throw new Exception(string.Format("ERROR Attribute '{0}' = '{1}' doesnt satisfy validation rule '{4}.{3}()'. ROW: {2}", field.Name, value, row, field.ValidationType, instance.GetType()));
                }
            }

        }
        public ValidationResult ValidateInstance(object objectInstance, ImportSettings loadSetup)
        {
            var result = new ValidationResult();
            if (loadSetup == null)
            {
                result.Status = false;
                result.Messages.Add("Import settings undefined");
                return result;
            }
            if (loadSetup.Regexs != null)
            {
                var regexs = loadSetup.Regexs.Where(x => x.ClassName == objectInstance.GetType().FullName).ToList();
                foreach (var regex in regexs)
                {
                    var field = objectInstance.GetType().GetProperties().FirstOrDefault(x => x.Name == regex.Attribute);
                    var value = field.GetValue(objectInstance).ToString();
                    if (regex != null && !string.IsNullOrEmpty(regex.Expression) && (Regex.Match(value, regex.Expression).Value == "" || string.IsNullOrEmpty(value)))
                    {
                        result.Status = false;
                        result.Messages.Add(string.Format("'{0}'='{1}' Doesnt satisfy REGEX:'{2}'", field.Name, value, regex.Expression));
                    }
                }
            }
            var entity = new Entity();
            if (loadSetup.Entities.Any(x => x.ClassName == objectInstance.GetType().FullName))
            {
                entity = loadSetup.Entities.FirstOrDefault(x => x.ClassName == objectInstance.GetType().FullName);
            }

            foreach (var field in entity.Fields.Where(x => x.Mandatory.ToLower() == "true"))
            {
                var fieldType = objectInstance.GetType().GetProperties().FirstOrDefault(x => x.Name == field.Attribute);
                if (fieldType == null)
                {
                    result.Status = false;
                    result.Messages.Add(string.Format("Attribute '{0}' not found in entity: '{1}'", field.Name, objectInstance.GetType().FullName));
                    return result;
                }
                var value = fieldType.GetValue(objectInstance).ToString();
                if (string.IsNullOrEmpty(value))
                {
                    result.Status = false;
                    result.Messages.Add(string.Format("'{0}' Is mandatory", field.Name));
                    return result;
                }
            }
            return result;
        }
        private static int WorksheetRowsNumber(ExcelWorksheet workSheet)
        {
            var worksheetRowsNumber = 3;
            for (var rowNumber = 3; rowNumber <= maxExcelRowsExpected; rowNumber++)
            {
                if (string.IsNullOrEmpty(workSheet.Cells[rowNumber, 1].Text))
                {
                    worksheetRowsNumber = rowNumber - 1;
                    break;
                }
            }

            return worksheetRowsNumber;
        }
    }
}