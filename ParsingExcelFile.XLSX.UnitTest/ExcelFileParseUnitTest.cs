using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ParsingExcelFile.XLSX.UnitTest
{
    [TestClass]
    public class ExcelFileParseUnitTest
    {
        [TestMethod]
        public void Add_AcceptTwoInt64Values_ReturnSumofTwoNums()
        {
            // Arrange
            Int64 x, y;
            x = 90;
            y = 89;
            ExcelFileParse obj = new ExcelFileParse();

            // Act
            Int64? result = null;

            result = obj.Add(x, y);

            Assert.AreNotEqual(null, result);
        }
    }
}
