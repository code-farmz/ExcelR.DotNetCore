using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using ExcelR.DotNetCore.Extensions;
using static ExcelR.DotNetCore.Enums.Excel;
using ExcelR.DotNetCore;

namespace ExcelRTest
{
    [TestClass]
    public class Test
    {
        [TestMethod]
        public string ExportExcel()
        {
            var filePath = $"{DocsDirctory}/test_{DateTime.Now.Ticks}.xlsx";
            var data = GetSampleData();
            data.ToExcel("Sheet1", Style.H3, Color.Aqua).Save(filePath);
            return filePath;
        }
        [TestMethod]
        public void ImportExcel()
        {
           var sheet = ExcelImporter.GetWorkSheet(ExportExcel());
           var data= ExcelImporter.Read<TestModel>(sheet);
        }


        [TestMethod]
        public string ExportCsv()
        {
            var filePath = $"{DocsDirctory}/test_{DateTime.Now.Ticks}.csv";
            var data = GetSampleData();
            data.ToCsv(filePath);
            return filePath;

        }

        [TestMethod]
        public void ImportCsv()
        {
            var data = CsvHelper.ReadFromFile<TestModel>(ExportCsv());
        }

        private List<TestModel> GetSampleData()
        {
            var list = new List<TestModel>
            {
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "Braat", LastName = "Lee"},
                new TestModel {IsMale = true,  FirstName = "Flintop"},
                new TestModel {IsMale = true, Dob = DateTime.Now.AddDays(15), FirstName = "Michel"},
                new TestModel {IsMale = true, Dob = DateTime.Now, FirstName = "Michel", LastName = "John"},
                new TestModel {IsMale = false, FirstName = "john", LastName = "Cena"}
            };
            return list;
        }

        #region private methods/props
        private string DocsDirctory
        {
            get
            {
                if (!Directory.Exists($"{Environment.CurrentDirectory}/docs"))
                    Directory.CreateDirectory($"{Environment.CurrentDirectory}/docs");
                return $"{Environment.CurrentDirectory}/docs";
            }
        }
        
        #endregion
    }
}
