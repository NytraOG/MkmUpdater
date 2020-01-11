using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Provider.Tests
{
    [TestClass]
    public class Unit
    {
        private XlWorkbookGenerator manager;

        [TestInitialize]
        public void Init()
        {
            manager = new XlWorkbookGenerator();
        }

        [TestMethod]
        public void GetExcelData_FileExists_CorrectAmountOfSheets()
        {
            //Arrange
            var expectedSheetCount = 4;

            //Act
            var actualSheetCount = manager.GetWorkbook().Sheets.Count;

            //Assert
            Assert.AreEqual(expectedSheetCount, actualSheetCount);
        }

        [TestMethod]
        public void GetExcelData_FileExists_CorrectlyMappedCreatureColumn()
        {
            //Arrange
            var expectedSheetName  = "Urza";
            var expectedColumnName = "Creature";

            //Act
            var actualWorkbook       = manager.GetWorkbook();
            var actualSheet          = actualWorkbook.Sheets.FirstOrDefault(x => x.Name       == expectedSheetName);
            var actualCreatureColumn = actualSheet?.Columns.FirstOrDefault(x => x.Description == expectedColumnName);

            //Assert
            Assert.IsNotNull(actualSheet);
            Assert.IsNotNull(actualCreatureColumn);
            Assert.IsTrue(actualCreatureColumn.Entrys.ContainsKey("Urza"));
            Assert.AreEqual(25, actualCreatureColumn.Entrys.First(e => e.Key == "Urza").Value);
            Assert.IsTrue(actualCreatureColumn.Entrys.ContainsKey("Phyrexian Metamorph"));
            Assert.AreEqual(5, actualCreatureColumn.Entrys.First(e => e.Key == "Phyrexian Metamorph").Value);
        }
    }
}