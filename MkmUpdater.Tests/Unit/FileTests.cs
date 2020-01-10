using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MkmUpdater.Tests.Unit
{
    [TestClass]
    public class FileTests
    {
        private Service1 serviceInstance;
        private string filePath;

        [TestInitialize]
        public void Init()
        {
            serviceInstance = new Service1();
            filePath = Constants.Path + "\\ServiceLog_" + DateTime.Now.DayOfWeek + ".txt";
        }

        [TestCleanup]
        public void CleanUp()
        {
            File.Delete(filePath);
        }

        [TestMethod]
        public void ChecksFilepath_ExistingPath()
        {
            //Arrange
            var timeString = DateTime.Now.ToShortTimeString();

            //Act
            serviceInstance.WriteToFile(timeString);
            var fileExists = File.Exists(filePath);

            //Assert
            Assert.IsTrue(fileExists);
        }
    }
}
