using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Punkstar.DocHelper.Xls2Object.SampleTestAssembly;

namespace Punkstar.DocHelper.Xls2Object.TestProject
{
    [TestClass]
    public class SetupTests
    {
        [TestMethod]
        public void ValidateBlankMainEntitySettingsCreation()
        {
            var instance = new MainEntity();
            var helper = new Xls2ObjectHelper();
            string blankMainEntityImportSettings = helper.CreateImportSettingsJsonByClass(instance, 2, null);
            var appDir = AppDomain.CurrentDomain.BaseDirectory;
            var blankMainEntityImportSettingsFileContent = File.ReadAllText(Path.Combine(appDir,"ImportSettingsFiles", "BlankMainEntityJsonSettingsSample.json"));
            Assert.AreEqual(blankMainEntityImportSettingsFileContent, blankMainEntityImportSettings);
        }
    }
}
