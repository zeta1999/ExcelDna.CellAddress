using System;
using ExcelDna;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CellAddressTests.AddIn
{
    [TestClass]
    public class CellAddressTests

    {
        [TestMethod]
        public void TestCreate() {
            var cell = new CellAddress("sheet1",0,0);
            Assert.AreEqual("$A$1", cell.LocalAddress);
            Assert.AreEqual("sheet1!$A$1", cell.FullAddress);

            cell = new CellAddress("Sheet 1",0,0);
            Assert.AreEqual("$A$1", cell.LocalAddress);
            Assert.AreEqual("'Sheet 1'!$A$1", cell.FullAddress);
        }

        [TestMethod]
        public void TestParse() {
            var address = CellAddress.Parse("A1");
            Assert.AreEqual(0, address.RowFirst);
            Assert.AreEqual(0, address.ColumnFirst);
            Assert.AreEqual(1, address.Count);

            address = CellAddress.Parse("Sheet 1!E:F");
            Assert.AreEqual(0, address.RowFirst);
            Assert.AreEqual(4, address.ColumnFirst);
            Console.WriteLine(address.FullAddress);

            address = CellAddress.Parse("Sheet1!9:9");
            Assert.AreEqual(8, address.RowFirst);
            Assert.AreEqual(0, address.ColumnFirst);
            Console.WriteLine(address.FullAddress);
        }

        [TestMethod]
        public void TestConstructFromRange() {
            var app = (Application)ExcelDnaUtil.Application;
            var sheet = (Worksheet) app.ActiveSheet;
            var range = sheet.Range["A1"];

            var address = new CellAddress(range);
            Assert.AreEqual("$A$1", address.LocalAddress);
        }

        [TestMethod]
        public void TestConstructFromEntireRow() {
            var app = (Application)ExcelDnaUtil.Application;
            var sheet = (Worksheet)app.ActiveSheet;
            var range = sheet.Range["A1"].EntireRow;

            var address = new CellAddress(range);
            Assert.AreEqual("$A", address.LocalAddress);
        }
    }

}
