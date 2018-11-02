using System;
using System.Net.Mime;
using ExcelDna;
using Xunit;

namespace Tests
{
    public class CellAddressTests {
        [Fact]
        public void TestCreate() {
            var cell = new CellAddress("sheet1", 0, 0);
            Assert.Equal("$A$1", cell.LocalAddress);
            Assert.Equal("sheet1!$A$1", cell.FullAddress);

            cell = new ExcelDna.CellAddress("Sheet 1", 0, 0);
            Assert.Equal("$A$1", cell.LocalAddress);
            Assert.Equal("'Sheet 1'!$A$1", cell.FullAddress);
        }

        [Fact]
        public void TestParse() {
            var address = CellAddress.Parse("A1");
            Assert.Equal(0, address.RowFirst);
            Assert.Equal(0, address.ColumnFirst);
            Assert.Equal(1, address.Count);

        }

        [Fact]
        public void TestParseEntireRow() {
            var address = CellAddress.Parse("Sheet 1!E:F");
            Assert.Equal(0, address.RowFirst);
            Assert.Equal(4, address.ColumnFirst);
            Console.WriteLine(address.FullAddress);

            address = CellAddress.Parse("Sheet1!9:9");
            Assert.Equal(8, address.RowFirst);
            Assert.Equal(0, address.ColumnFirst);
            Console.WriteLine(address.FullAddress);

        }
    }
}
