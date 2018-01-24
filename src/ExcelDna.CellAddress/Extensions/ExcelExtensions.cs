using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna.Extensions {
    /// <summary>
    /// Excel App 扩展方法
    /// </summary>
    internal static class ExcelApp {
        /// <summary>
        /// Excel Application 调用句柄
        /// </summary>
        /// <param name="application"></param>
        public delegate void ApplicationInvokeAction(Application application);

        public delegate TValue ApplicationFunc<out TValue>(Application application);


        /// <summary>
        ///     执行同步命令
        /// </summary>
        /// <param name="action"></param>
        public static void Excute(this ApplicationInvokeAction action) {
            Application xlApp = null;
            try {
                xlApp = ExcelDnaUtil.Application as Application;
                if (xlApp == null) {
                    throw new InvalidOperationException("Application is Null");
                }
                xlApp.ScreenUpdating = false;

                action(xlApp);
            } catch (InvalidOperationException ioe) {
                //当前 ExcelApplication 不可用
                Debug.Print(ioe.Message);
            } catch (Exception ex) {
                Debug.Print(ex.Message);
            } finally {
                try {
                    if (xlApp != null) {
                        xlApp.ScreenUpdating = true;
                        xlApp.EnableEvents = true;
                    }
                } catch (COMException) {
                    
                }
            }
        }


        public static TValue Return<TValue>(ApplicationFunc<TValue> func) {
            object xlApp = null;
            try {
                xlApp = ExcelDnaUtil.Application;
                var application = xlApp as Application;
                return func(application);
            } catch (InvalidOperationException ioe) {
                //当前 ExcelApplication 不可用
                Debug.Print("ExcelApp.Return<TValue> error:" + ioe.Message);
                return default(TValue);
            } finally {
                if (xlApp != null) {
                    //Marshal.ReleaseComObject(xlApp);
                }
            }
        }
    }
}