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
        List<EntityRange> EntitiesRanges;
        public LoadSetup setup;
        public Assembly assembly;
        public void AssignValue(object instance, PropertyInfo prop, Field field, EntityRange range, string value, int row)
        {
            try
            {
                switch (field.FieldType.ToLower())
                {
                    case "string":
                        if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna '{0}' con el valor para el atributo '{1}' (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
                        if (string.IsNullOrEmpty(value))
                        {
                            prop.SetValue(instance, "", null);
                            return;
                        }

                        if (setup.Regexs != null)
                        {
                            var regex = setup.Regexs.FirstOrDefault(x => x.Attribute == field.Attribute && x.ClassName == range.ClassName) != null ? setup.Regexs.FirstOrDefault(x => x.Attribute == field.Attribute && x.ClassName == range.ClassName).Expression : "";
                            if (regex != null && !string.IsNullOrEmpty(regex) && (Regex.Match(value, regex).Value == "" || string.IsNullOrEmpty(value)))
                                throw new Exception(string.Format("ERROR(Linea {2}) Entidad '{5}': Columna '{0}' con el valor para el atributo '{1}' (tipo: '{3}') no cumple con la expresión regular '{4}'", field.Name, field.Attribute, row, field.FieldType, regex, range.Name));
                        }
                        prop.SetValue(instance, value, null);
                        break;
                    case "int":
                    case "int32":
                    case "int64":
                        if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad '{4}': Columna '{0}' con el valor para el atributo '{1}' (tipo: '{3}') es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
                        if (string.IsNullOrEmpty(value))
                        {
                            try
                            {
                                prop.SetValue(instance, null, null);
                            }
                            catch (Exception ex)
                            {
                                prop.SetValue(instance, 0, null);
                            }
                            return;
                        }
                        int n;
                        bool isNumeric = int.TryParse(value, out n);
                        if (!isNumeric)
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
                        prop.SetValue(instance, n, null);
                        break;
                    case "datetime":
                        if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
                        if (string.IsNullOrEmpty(value))
                        {
                            prop.SetValue(instance, new DateTime(), null);
                            return;
                        }

                        DateTime date;
                        if (!DateTime.TryParse(value, out date))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
                        prop.SetValue(instance, date, null);
                        break;
                    case "bool":
                    case "boolean":
                        if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
                        if (string.IsNullOrEmpty(value))
                        {
                            prop.SetValue(instance, false, null);
                            return;
                        }
                        bool varBool;
                        if (!bool.TryParse(value, out varBool))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
                        prop.SetValue(instance, varBool, null);
                        break;
                    case "guid":
                        if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
                        if (string.IsNullOrEmpty(value))
                            return;
                        Guid varGuid;
                        if (!Guid.TryParse(value, out varGuid))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} no cumple con el tipo {3}", field.Name, field.Attribute, row, field.FieldType.ToUpper(), range.Name));
                        prop.SetValue(instance, varGuid, null);
                        break;
                    default:
                        if (field.Mandatory == "true" && string.IsNullOrEmpty(value))
                            throw new Exception(string.Format("ERROR(Linea {2}) Entidad {4}: Columna {0} con el valor para el atributo {1} (tipo: {3}) es obligatorio", field.Name, field.Attribute, row, field.FieldType, range.Name));
                        if (string.IsNullOrEmpty(value))
                            return;
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
                            catch (Exception ex)
                            {
                                parseDirecto = false;
                            }
                            if (!parseDirecto)
                                try
                                {
                                    var encontrado = false;
                                    foreach (var declaredField in ((TypeInfo)childType).DeclaredFields)
                                        foreach (var customAttribute in declaredField.CustomAttributes)
                                        {
                                            if (customAttribute.ConstructorArguments != null && customAttribute.ConstructorArguments.Count > 0)
                                                if (customAttribute.ConstructorArguments.Any(x => x.Value.ToString().ToLower() == value.ToLower()))
                                                {
                                                    childInstance = Enum.Parse(childType, declaredField.Name);
                                                    prop.SetValue(instance, childInstance, null);
                                                    encontrado = true;
                                                    break;
                                                }
                                            if (encontrado)
                                                break;
                                        }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception(string.Format("No ha sido posible interretar el valor '{0}' en el enumerador '{1}.Error:{2}'", value, field.FieldType, ex.Message));
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
                        break;
                }




            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        public LoadSetup CreateLoadSetupByObject(object instance, int deepLevel, string[] excludedClasses)
        {
            var output = new LoadSetup();
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
        public LoadSetup GetSetupFile(Stream JSONFile)
        {
            LoadSetup loadSetup = null;

            using (var reader = new StreamReader(JSONFile))
            {
                StreamReader readStream = new StreamReader(JSONFile, Encoding.UTF8);
                string jsonString = "";
                jsonString = jsonString + readStream.ReadToEnd();
                jsonString = jsonString.Replace(@"\", " ");
                loadSetup = JsonConvert.DeserializeObject<LoadSetup>(jsonString);
            }
            return loadSetup;
        } 
        public List<EntityRange> GetEntityRanges(List<Entity> Entities, ExcelPackage XLSFile)
        {
            try
            {
                var entityRanges = new List<EntityRange>();
                var currentSheet = XLSFile.Workbook.Worksheets;
                foreach (var entity in Entities)
                {
                    if (entity.SpreadSheetName == null && entity.Parent == null) continue;
                    var workSheet = currentSheet.FirstOrDefault(x => x.Name == entity.SpreadSheetName);
                    if (workSheet == null)
                        throw new Exception(string.Format("Worksheet '{0}' not found.", entity.SpreadSheetName));
                    var noOfCol = workSheet.Dimension.End.Column;
                    var lastFoundEntity = "";
                     for (int i = 1; i <= noOfCol; i++)
                    {
                        if (workSheet.Cells[1, i].Text != "")
                        {
                            if (lastFoundEntity == "")
                            {
                                lastFoundEntity = workSheet.Cells[1, i].Text;
                                entityRanges.Add(new EntityRange { SpreadSheetName = entity.SpreadSheetName, Name = lastFoundEntity, Start = i, ClassName = Entities.Any(x => x.Name == lastFoundEntity) ? Entities.FirstOrDefault(x => x.Name == lastFoundEntity).ClassName : "NoName" });
                            }
                            else if (lastFoundEntity != workSheet.Cells[1, i].Text)
                            {
                                lastFoundEntity = workSheet.Cells[1, i].Text;
                                entityRanges.Last().End = i - 1;
                            }
                        }
                    }
                    entityRanges.Last().End = noOfCol;
                }
                if (entityRanges.Count == 0)
                {
                    entityRanges = new List<EntityRange>();
                    currentSheet = XLSFile.Workbook.Worksheets;
                     foreach (var entity in Entities.FirstOrDefault().Entities.Where(x => x.SpreadSheetName != null))
                    {
                        var workSheet = currentSheet.FirstOrDefault(x => x.Name == entity.SpreadSheetName);
                        if (workSheet == null)
                            throw new Exception(string.Format("Worksheet '{0}' not found.", entity.SpreadSheetName));
                        var noOfCol = workSheet.Dimension.End.Column;
                        var lastFoundEntity = "";
                        for (int i = 1; i <= noOfCol; i++)
                        {
                            if (workSheet.Cells[1, i].Text != "")
                            {
                                if (lastFoundEntity == "")
                                {
                                    lastFoundEntity = workSheet.Cells[1, i].Text;
                                    entityRanges.Add(new EntityRange { SpreadSheetName = entity.SpreadSheetName, Name = lastFoundEntity, Start = i, ClassName = Entities.FirstOrDefault().Entities.FirstOrDefault(x => x.Name == lastFoundEntity).ClassName });
                                }
                                else if (lastFoundEntity != workSheet.Cells[1, i].Text)
                                {
                                    lastFoundEntity = workSheet.Cells[1, i].Text;
                                    entityRanges.Last().End = i - 1;
                                }
                            }
                        }
                        entityRanges.Last().End = noOfCol;
                    }

                }
                var subEntityRanges = new List<EntityRange>();
                foreach (var entityRange in entityRanges)
                {
                    var entity = Entities.FirstOrDefault(x => x.Name == entityRange.Name);
                    if (entity != null)
                    {
                        var subEntities = entity.Entities.Where(x => !string.IsNullOrEmpty(x.SpreadSheetName)).ToList();
                        if (subEntities != null && subEntities.Count > 0)
                        {
                            subEntityRanges.AddRange(GetEntityRanges(subEntities, XLSFile));

                        }
                    }

                }
                entityRanges.AddRange(subEntityRanges);
                foreach (var entity in Entities.FirstOrDefault().Entities.Where(x => !string.IsNullOrEmpty(x.SpreadSheetName)).ToList())
                {
                    if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.SpreadSheetName)))
                        entityRanges.AddRange(GetEntityRanges(entity.Entities.Where(x => !string.IsNullOrEmpty(x.SpreadSheetName)).ToList(), XLSFile));
                }
                return entityRanges;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        public bool Evaluate(Condition condition, object instance)
        {
            if (string.IsNullOrEmpty(condition.Entity))
            {
                var properties = instance.GetType().GetProperties();
                var prop = instance.GetType().GetProperty(condition.Field);
                if (condition.Operation == Enums.Operator.Equal)
                    return prop.GetValue(instance, null).ToString() == condition.Value;
                if (condition.Operation == Enums.Operator.GreaterThan)
                    return double.Parse(prop.GetValue(instance, null).ToString()) > double.Parse(condition.Value);
                if (condition.Operation == Enums.Operator.LesserThan)
                    return double.Parse(prop.GetValue(instance, null).ToString()) < double.Parse(condition.Value);
            }
            return false;
        }
        public string GetCellValue(ExcelWorksheet workSheet, EntityRange range, Field field, int row)
        {
            for (var i = range.Start; i <= range.End; i++)
                if (workSheet.Cells[2, i].Text.ToLower() == field.Name.ToLower())
                {
                    return workSheet.Cells[row, i].Text;
                }
            return "";
        }
        public string GetCellValue(ExcelWorksheet workSheet, EntityRange range, string columnName, int row)
        {
            for (var i = range.Start; i <= range.End; i++)
                if (workSheet.Cells[2, i].Text.ToLower() == columnName.ToLower())
                {
                    return workSheet.Cells[row, i].Text;
                }
            return "";
        }
        public List<Entity> GetDependantEntities(object instance, int deep, int maxDeep, string[] excludedClasses)
        {
            var entities = new List<Entity>();
            var instanceType = instance.GetType();
            var fields = new List<Field>();
            foreach (var property in instanceType.GetProperties())
            {
                if (!property.CanWrite || !property.GetSetMethod(true).IsPublic)
                    continue;
                var excluded = false;
                foreach (var excludedClass in excludedClasses)
                {
                    if (property.PropertyType.FullName.Contains(excludedClass))
                    {
                        excluded = true;
                        break;
                    }
                }
                if (excluded) continue;
                string propertyTypeName = property.PropertyType.Name.ToLower();
                if (propertyTypeName.Contains("nullable"))
                    propertyTypeName = property.PropertyType.GenericTypeArguments[0].Name.ToLower();
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

                        if (property.PropertyType.Name.Contains("List"))
                        {
                            var listType = typeof(List<>);
                            var constructedListType = listType.MakeGenericType(property.PropertyType.GenericTypeArguments[0]);
                            var listInstance = Activator.CreateInstance(constructedListType);
                            var dependantEntities = new List<Entity>();
                            if (listInstance.GetType().GenericTypeArguments.Count() > 0)
                            {
                                var testObject = Activator.CreateInstance(listInstance.GetType().GenericTypeArguments[0]);
                                if (deep < maxDeep)
                                    dependantEntities = GetDependantEntities(testObject, deep + 1, maxDeep, excludedClasses);
                            }

                            var entity = GetMainEntity(listInstance, excludedClasses);
                            entity.Name = property.Name;
                            entity.ParentAttribute = property.Name;
                            entity.IsList = true;
                            entity.ExcelLookUpField = string.Format("'{0}' Excel column lookup ", property.Name);
                            entity.ParentLookUpField = "field in parent to look up for";
                            entity.Entities = dependantEntities;
                            entities.Add(entity);
                        }
                        else
                        {
                            var propertyInstance = Activator.CreateInstance(property.PropertyType);
                            var propertyEntity = GetMainEntity(propertyInstance, excludedClasses);
                            propertyEntity.Name = property.Name;
                            propertyEntity.ParentAttribute = property.Name;
                            propertyEntity.ExcelLookUpField = string.Format("'{0}' Excel column lookup ", property.Name);
                            propertyEntity.ParentLookUpField = "field in parent to look up for";
                            propertyEntity.IsList = false;
                            propertyEntity.Parent = instanceType.FullName;
                            var propertyDependantEntities = new List<Entity>();
                            if (deep + 1 < maxDeep)
                            {
                                propertyDependantEntities = GetDependantEntities(propertyInstance, deep + 1, maxDeep, excludedClasses);
                                foreach (var propertyDependantEntity in propertyDependantEntities)
                                {
                                    propertyDependantEntity.Parent = property.Name;

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
        public List<Field> GetEntityFields(object instance, string[] excludedClasses)
        {
            var instanceType = instance.GetType();
            if (instanceType.Name.Contains("List"))
                instanceType = instance.GetType().GenericTypeArguments[0];

            var fields = new List<Field>();
            foreach (var property in instanceType.GetProperties())
            {
                switch (property.PropertyType.Name.ToLower())
                {
                    case "string":
                    case "int":
                    case "int32":
                    case "int64":
                    case "datetime":
                    case "bool":
                    case "boolean":
                    case "guid":
                        MethodInfo setMethod = property.GetSetMethod();
                        if (setMethod != null)
                        {
                            var excluded = false;
                            foreach (var excludedClass in excludedClasses)
                            {
                                if (property.PropertyType.FullName.Contains(excludedClass))
                                {
                                    excluded = true;
                                    break;
                                }
                            }
                            if (excluded) continue;

                            var field = new Field();
                            field.Attribute = property.Name;
                            field.FieldType = property.PropertyType.Name;
                            field.Mandatory = "false";
                            field.Name = "Excel column name";
                            field.ValidationType = "";
                            fields.Add(field);
                        }
                        break;
                }
            }
            return fields;
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
                if (entity.SpreadSheetName == null && entity.Parent == null)
                {
                    var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                    var parentInstance = Activator.CreateInstance(type);
                    if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.SpreadSheetName)))
                        list.AddRange(GetObjectsFromExcel(parentInstance, XLSFile, entity.Entities.Where(x => !string.IsNullOrEmpty(x.SpreadSheetName)).ToList(), EntityRanges));
                    continue;
                }
                var workSheet = currentSheet.FirstOrDefault(x => x.Name == entity.SpreadSheetName);
                var totalSpreadsheetRows = 3;
                for (var i = 3; i <= 100000; i++)
                    if (string.IsNullOrEmpty(workSheet.Cells[i, 1].Text))
                    {
                        totalSpreadsheetRows = i - 1;
                        break;
                    }
                for (int currentRowNumber = 3; currentRowNumber <= totalSpreadsheetRows; currentRowNumber++)
                {
                    var range = EntitiesRanges.FirstOrDefault(x => x.Name == entity.Name);
                    if (range == null)
                        throw new Exception(string.Format("No se ha encontrado un rango para la entidad '{0}' en la hoja '{1}' en la sección '{2}'", entity.ClassName, entity.SpreadSheetName, entity.Name));
                    var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                    var instance = Activator.CreateInstance(type);
                    PopulateInstance(instance, Entities.FirstOrDefault(x => x.Name == entity.Name), workSheet, range, currentRowNumber);
                    if (Entities.FirstOrDefault(x => x.Name == entity.Name).Entities.Any(x => !string.IsNullOrEmpty(x.SpreadSheetName)))
                    {
                        instance = GetObjectsFromExcel(instance, XLSFile, Entities.FirstOrDefault(x => x.Name == entity.Name).Entities.Where(x => !string.IsNullOrEmpty(x.SpreadSheetName)).ToList(), EntitiesRanges).FirstOrDefault();
                    }
                    PropertyInfo[] propsList = parentObject.GetType().GetProperties();
                    var excelLookupValue = GetCellValue(workSheet, range, entity.ExcelLookUpField, currentRowNumber);
                    var parentType = parentObject.GetType();
                    var parentLookupProperty = parentType.GetProperty(entity.ParentLookUpField);
                    if (entity.ExcelLookUpField != "" && !entity.ExcelLookUpField.Contains(" Excel column lookup") && string.IsNullOrWhiteSpace(excelLookupValue))
                        throw new Exception(string.Format("La celda '{0}' no contiene valor o no existe en la hoja {1}", entity.ExcelLookUpField, entity.SpreadSheetName));
                    var prop = propsList.FirstOrDefault(x => x.Name.Equals(entity.ParentAttribute));
                    if (prop == null)
                        throw new Exception(string.Format("'{0}' is not an attribute of '{1}' class", entity.ParentAttribute, entity.Parent));
                    if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition().Name.Contains("List"))
                    {
                        if (parentLookupProperty != null && parentLookupProperty.Name != "")
                        {
                            if (parentObject.GetType().GetProperty(parentLookupProperty.Name).GetValue(parentObject, null) == excelLookupValue)
                            {
                                IList reflectedList = (IList)prop.GetValue(parentObject);
                                reflectedList.Add(instance);
                                parentObject.GetType().GetProperty(entity.ParentAttribute).SetValue(parentObject, reflectedList);
                            }
                        }
                        else
                        {
                            IList reflectedList = (IList)prop.GetValue(parentObject);
                            reflectedList.Add(instance);
                            parentObject.GetType().GetProperty(entity.ParentAttribute).SetValue(parentObject, reflectedList);
                        }

                    }
                    else if (parentLookupProperty != null && parentLookupProperty.Name != "")
                    {
                        if (parentObject.GetType().GetProperty(parentLookupProperty.Name).GetValue(parentObject, null) == excelLookupValue)
                        {
                            prop.SetValue(parentObject, instance, null);
                        }
                    }
                    else
                        prop.SetValue(parentObject, instance, null);

                }
            }
            list.Add(parentObject);
            return list;

        }
        public List<dynamic> GetObjectsFromExcel(ExcelPackage XLSFile, Stream JSONFile)
        {
            var list = new List<dynamic>();
            setup = GetSetupFile(JSONFile);

            
            var entities = new List<Entity>();

            entities = setup.Entities;
           
            EntitiesRanges = GetEntityRanges(entities, XLSFile);
            foreach (var entity in entities)
            {
                var type = assembly.GetTypes().FirstOrDefault(x => x.FullName == entity.ClassName);
                var parentInstance = Activator.CreateInstance(type);
                if (entity.Entities.Any(x => !string.IsNullOrEmpty(x.SpreadSheetName)))
                    list.AddRange(GetObjectsFromExcel(parentInstance, XLSFile, entity.Entities.Where(x => !string.IsNullOrEmpty(x.SpreadSheetName)).ToList(), EntitiesRanges));
            }

            return list;
        }
        public string GetLoadSetupJsonByClass(object instance, int deepLevel, string[] excludedClasses)
        {
            var output = CreateLoadSetupByObject(instance, deepLevel, excludedClasses);
            return JsonConvert.SerializeObject(output);
        }
        public void LoadAssembly(Stream inputStream)
        {
            if (inputStream == null) return;
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
                LoadAssembly(inputStream);
            else
                assembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(x => x.ManifestModule.ScopeName == assemblyName);
        }
        public object PopulateAttribute(object instance, ExcelWorksheet workSheet, EntityRange range, Field field, int row)
        {
            var prop = instance.GetType().GetProperties().FirstOrDefault(x => x.Name.Equals(field.Attribute));
            if (prop == null)
                throw new Exception(string.Format("El atributo {0} no encontrado en la clase {1}. Verifique el assembly", field.Attribute, instance.GetType()));
            if (null != prop && prop.CanWrite)
            {
                var value = GetCellValue(workSheet, range, field, row);
                AssignValue(instance, prop, field, range, value, row);
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
                        foreach (var _Entity in _ConditionalEntity.Entities)
                        {
                            var conditionalRange = EntitiesRanges.FirstOrDefault(x => x.Name == _Entity.Name);
                            var conditionalType = assembly.GetTypes().FirstOrDefault(x => x.FullName == _Entity.ClassName);
                            var conditionalInstance = Activator.CreateInstance(conditionalType);
                            var childInstance = PopulateInstance(conditionalInstance, _Entity, workSheet, conditionalRange, row);
                            var prop = instance.GetType().GetProperty(_ConditionalEntity.Attribute, BindingFlags.Public | BindingFlags.Instance);
                            if (prop == null)
                                throw new Exception(string.Format("El atributo {0} no encontrado en la clase {1}. Verifique el assembly", _ConditionalEntity.Attribute, conditionalInstance.GetType()));
                            prop.SetValue(instance, childInstance, null);
                        }
                }
            ValidateEntity(entity, instance, row);
            return instance;
        }
        public object PopulateObjectWithRandomData(object objectInstance)
        {
            if (objectInstance == null)
                return null;
            foreach (var field in objectInstance.GetType().GetProperties())
            {
                if (!field.CanWrite || !field.GetSetMethod(true).IsPublic) continue;

                string propertyTypeName = field.PropertyType.Name.ToLower();
                if (propertyTypeName.Contains("nullable"))
                    propertyTypeName = field.PropertyType.GenericTypeArguments[0].Name.ToLower();
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
                        field.SetValue(objectInstance, (Int32)random.Next(1, 10000), null);
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
                                catch (Exception ex)
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
                                field.SetValue(objectInstance, childobjectInstance.GetType().GetEnumValues().GetValue(0), null);
                            else
                                field.SetValue(objectInstance, childobjectInstance, null);
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
        public void ValidateEntity(Entity entity, object instance, int row)
        {
            if (!string.IsNullOrEmpty(entity.ValidationType))
            {
                MethodInfo method = instance.GetType().GetMethod(entity.ValidationType);
                if (method == null)
                    throw new Exception(string.Format("Método de validación {0} no encontrado en la clase {1}. Verifique el assembly", entity.ValidationType, entity.ClassName));
                object result = method.Invoke(instance, null);
                if (!((bool)result))
                    throw new Exception(string.Format("ERROR El registro en la fila {0} no cumple con la validación{1}", row, entity.ValidationType));
            }
        }
        public void ValidateField(object instance, Field field, string value, int row)
        {
            if (!string.IsNullOrEmpty(field.ValidationType))
            {
                MethodInfo method = instance.GetType().GetMethod(field.ValidationType);
                if (method == null)
                    throw new Exception(string.Format("Método de validación {0} no encontrado en la clase {1}. Verifique el assembly", field.ValidationType, instance.GetType()));
                object result = method.Invoke(instance, null);
                if (!((bool)result))
                    throw new Exception(string.Format("ERROR El campo {0} con el valor '{1}' no cumple con la validación '{4}.{3}()' en la fila {2}", field.Name, value, row, field.ValidationType, instance.GetType()));
            }

        }
        public ValidationResult ValidateInstance(object objectInstance, LoadSetup loadSetup)
        {
            var result = new ValidationResult();
            if (loadSetup == null)
            {
                result.Status = false;
                result.Messages.Add("configuración de carga no definida");
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
                        result.Messages.Add(string.Format("'{0}'='{1}' no cumple con la expresión regular:'{2}'", field.Name, value, regex.Expression));
                    }
                }
            }
            var entity = new Entity();
            if (loadSetup.Entities.Any(x => x.ClassName == objectInstance.GetType().FullName))
                entity = loadSetup.Entities.FirstOrDefault(x => x.ClassName == objectInstance.GetType().FullName);
            foreach (var field in entity.Fields.Where(x => x.Mandatory.ToLower() == "true"))
            {
                var fieldType = objectInstance.GetType().GetProperties().FirstOrDefault(x => x.Name == field.Attribute);
                if (fieldType == null)
                {
                    result.Status = false;
                    result.Messages.Add(string.Format("{0} No encontrado en la entidad {1}", field.Name, objectInstance.GetType().FullName));
                    return result;
                }
                var value = fieldType.GetValue(objectInstance).ToString();
                if (string.IsNullOrEmpty(value))
                {
                    result.Status = false;
                    result.Messages.Add(string.Format("{0} Es obligatorio", field.Name));
                    return result;
                }
            }
            return result;
        }
    }
}