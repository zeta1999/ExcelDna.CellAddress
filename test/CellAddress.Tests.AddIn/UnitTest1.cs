using ExcelDna;
using ExcelDna.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CellAddressTests.AddIn
{
    [TestClass]
    public class CellAddressTest

    {
        [TestMethod]
        public void Test1() {
            var cell = new CellAddress("sheet1",0,0);
            Assert.AreEqual("A1", cell.LocalAddress);
            Assert.AreEqual("sheet1!A1", cell.FullAddress);

            cell = new CellAddress("Sheet 1",0,0);
            Assert.AreEqual("A1", cell.LocalAddress);
            Assert.AreEqual("'Sheet 1'!A1", cell.FullAddress);
        }

        [TestMethod]
        public void TestConstruct1() {
            var cell = new CellAddress(new ExcelReference(0,0));
            Assert.AreEqual("A1",cell.LocalAddress);
        }
    }
}
