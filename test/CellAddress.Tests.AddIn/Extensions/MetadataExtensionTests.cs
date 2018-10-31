using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CellAddressTests.AddIn.Extensions {

    [TestClass]
    public class MetadataExtensionTests {

        class MyClass {
            
            [DefaultValue("Test 1")]
            public string StringWithDefaultValue { get; set; }

            public int IntNoDefault { get; set; }
        }

        [TestMethod]
        public void TestGetDefaultValue() {
            var instance = new MyClass();
            var value  = instance.GetDefaultValue(a => a.StringWithDefaultValue);
            Assert.AreEqual("Test 1",value);

            Assert.AreEqual(0,instance.GetDefaultValue(a=>a.IntNoDefault));
        }
    }
}
