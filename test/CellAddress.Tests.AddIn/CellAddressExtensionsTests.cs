using ExcelDna;
using ExcelDna.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CellAddressTests.AddIn {
    [TestClass]
    public class CellAddressExtensionsTests
    {
        [TestMethod]
        public void TestClearContents() {
            const string msg = "Test Clear Contents";
            var cell = CellAddress.Parse("A1");
            cell.SetValue(msg);
            Assert.AreEqual(msg, cell.GetValue<string>());

            cell.ClearContents();
            Assert.IsTrue(cell.GetValue<object>().IsNull());
        }

        [TestMethod]
        public void TestGetCellWithIndex() {
            var cells = CellAddress.Parse("A1:F5");
            var cell = cells.GetCell(1,XlFillDirection.ColumnFirst);
            Assert.AreEqual("A2", cell.LocalAddress);

            var cell1 = cells.GetCell(2, XlFillDirection.RowFirst);
            Assert.AreEqual("C1",cell1.LocalAddress);
        }
    }
}
