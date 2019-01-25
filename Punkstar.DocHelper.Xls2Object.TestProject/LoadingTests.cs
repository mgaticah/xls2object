using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Punkstar.DocHelper.Xls2Object.SampleTestAssembly;

namespace Punkstar.DocHelper.Xls2Object.TestProject
{
    [TestClass]
    public class LoadingTests
    {
        const string sampleAssemblyNAme = "Punkstar.DocHelper.Xls2Object.SampleTestAssembly.dll";
        const string AssemblySampleFolder = "AssemblySample";
        const string ExcelSamplesFolder = "ExcelSampleFiles";
        const string ExcelsampleOneName = "ExcelSampleOne.xlsx";
        const string ImportSettingsFolder = "ImportSettingsFiles";
        const string ExcelSampleOneMainEntitySettingsName = "ExcelSampleOneMainEntitySettings.json";
        [TestMethod]
        public void InstanceEntitiesFromExcelSampleOne()
        {
            var mainEntityInstance = new MainEntity();
            var helper = new Xls2ObjectHelper();
            var appDir = AppDomain.CurrentDomain.BaseDirectory;
            var assemblyFileName = Path.Combine(appDir, AssemblySampleFolder, sampleAssemblyNAme);
            var assemblyStream = File.Open(assemblyFileName, FileMode.Open, FileAccess.Read);
            helper.LoadAssembly(assemblyStream);
            var excelSampleOneStream = File.Open(Path.Combine(appDir, ExcelSamplesFolder, ExcelsampleOneName), FileMode.Open);
            var importSettingsFileName = Path.Combine(appDir, ImportSettingsFolder, ExcelSampleOneMainEntitySettingsName);
            var importSettingsFileStream = File.Open(importSettingsFileName, FileMode.Open, FileAccess.Read);
            var list = helper.GetObjectsFromExcel(excelSampleOneStream, importSettingsFileStream);
            var castedList = new List<MainEntity>();
            foreach (var entity in list)
            {
                var castedEntity = new MainEntity();
                castedEntity.MainEntityId = entity.MainEntityId;
                castedEntity.StringFieldSample= entity.StringFieldSample;
                castedEntity.DateTimeFieldSample= entity.DateTimeFieldSample;
                castedList.Add(castedEntity);
            }
            excelSampleOneStream.Close();
            importSettingsFileStream.Close();
            assemblyStream.Close();
            Assert.AreEqual(castedList.Count, 3);
            Assert.AreEqual(castedList[2].StringFieldSample, "you need to load");
        }
    }
}
