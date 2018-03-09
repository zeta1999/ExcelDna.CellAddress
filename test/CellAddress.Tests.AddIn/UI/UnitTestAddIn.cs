using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace CellAddressTests.AddIn.UI {
    public class UnitTestAddIn : IExcelAddIn {
        #region Implementation of IExcelAddIn

        public void AutoOpen() {
            Trace.Listeners.Add(new ConsoleTraceListener());
        }

        public void AutoClose() {
            
        }

        #endregion
    }
}
