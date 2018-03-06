using System;
using Xunit;
using ExcelDna;
using ExcelDna.Integration;

namespace ExcelDna.Tests
{
    public class CellAddressTest

    {
        [Fact]
        public void Test1() {
            var cell = new CellAddress("sheet1",0,0);
            Assert.Equal("A1", cell.LocalAddress);
            Assert.Equal("sheet1!A1", cell.FullAddress);

            cell = new CellAddress("Sheet 1",0,0);
            Assert.Equal("A1", cell.LocalAddress);
            Assert.Equal("'Sheet 1'!A1", cell.FullAddress);
        }

        [Fact]
        public void TestConstruct1() {
            var cell = new CellAddress(new ExcelReference(0,0));
            Assert.Equal("A1",cell.LocalAddress);
        }
    }
}
