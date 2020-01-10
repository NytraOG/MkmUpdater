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
        public void GetWorkbook_CorrectPath_WorkbookInitialized()
        {
            //Arrange & Act
            var workbook = manager.GetWorkbook();

            //Assert
            Assert.IsNotNull(workbook);
        }

        [TestMethod]
        public void GetSheets_CorrectPath_RightAmountOfSheets()
        {
            //Arrange
            var expectedSheetsCount = 4;

            //Act
            var actual = manager.GetSheets();

            //Assert
            Assert.AreEqual(expectedSheetsCount, actual.Count);
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
