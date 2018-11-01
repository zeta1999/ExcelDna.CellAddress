using ExcelDna;
using ExcelDna.Extensions;
using Xunit;

namespace Tests {
    public class TestExtensions {


        [Fact]
        public void TestMax() {
            var cell1 = CellAddress.Parse("A1");
            var cell2 = CellAddress.Parse("C4");
            var max = cell1.Max(cell2);
            Assert.Equal(max,cell2);

            max = new[] {cell1, cell2}.Max();
            Assert.Equal(max, cell2); 
        }

        [Fact]
        public void TestMin() {
            var cell1 = CellAddress.Parse("A1");
            var cell2 = CellAddress.Parse("C4");
            var min = cell1.Min(cell2);
            Assert.Equal(min, cell1);

            min = new[] {cell1, cell2}.Min();
            Assert.Equal(min, cell1);
        }

        [Fact]
        public void TestOffset() {
            var cell1 = CellAddress.Parse("A1");
            var offset1 = cell1.Offset();
            Assert.Equal(offset1, cell1);

            var offset2 = cell1.Offset(1, 1);
            Assert.Equal("$B$2",offset2.LocalAddress);

        }

        [Fact]
        public void TestGetRange() {
            var cell1 = CellAddress.Parse("A1");
            var cell2 = CellAddress.Parse("C4");
            var range = new[] {cell1, cell2}.GetRange();
            Assert.Equal("$A$1:$C$4", range.LocalAddress);

        }
    }
}