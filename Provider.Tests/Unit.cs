using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Provider.Tests
{
    [TestClass]
    public class Unit
    {
        private ExcelManager manager;

        [TestInitialize]
        public void Init()
        {
            manager = new ExcelManager();
        }

        [TestMethod]
        public void GetExcelData_FileExists_CorrectAmountOfSheets()
        {
            //Arrange
            var expectedSheetCount = 4;

            //Act
            var actualSheetCount = manager.GetExcelData().Sheets.Count;

            //Assert
            Assert.AreEqual(expectedSheetCount, actualSheetCount);
        }
    }
}
