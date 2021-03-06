﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CellAddressTests.AddIn.UI {
    public static class CellAddRressExtensionsTests {
        private const string UnitTestMenuName = "测试 CellAddress ";

        [ExcelCommand(Description = "测试全部", MenuName = UnitTestMenuName, MenuText = "测试全部")]
        public static void TestAll() {
            var provider = new UnitTestProvider();
            provider.Resolve();
            foreach (var category in provider.TestCategories) {
                Trace.TraceInformation("Test {0}", category.Name);
                foreach (var method in category.TestMethods) {
                    try {
                        Trace.TraceInformation("Invoke Method {0}", method.Name);
                        method.Invoke(category.Instance, new object[0]);
                    } catch (Exception ex) {
                        Trace.TraceWarning("invoke method {0} failed,{1}", method.Name, ex);
                        MessageBox.Show($"invoke method {method.Name} failed,{ex.ToString()}");
                    }
                } 
            }
        }

        private static IEnumerable<MethodInfo> GetMethods(Type testType) {
            var methods = typeof(CellAddRressExtensionsTests).GetMethods(BindingFlags.Instance | BindingFlags.Public);
            return methods.Where(IsUnitTestMethod);
        }

        private static bool IsUnitTestMethod(MethodInfo method) {
            return method.GetParameters().Length == 0 && method.GetCustomAttributes(typeof(TestMethodAttribute)).Any();
        }
    }

    public class TestCategory {

        public TestCategory(Type type) {
            this.UnitTesType = type;
            this.TestMethods = GetMethods(type).ToArray();
        }

        private static IEnumerable<MethodInfo> GetMethods(Type testType) {
            var methods = testType.GetMethods(BindingFlags.Instance | BindingFlags.Public);
            return methods.Where(IsUnitTestMethod);
        }

        private static bool IsUnitTestMethod(MethodInfo method) {
            return method.GetParameters().Length == 0 && method.GetCustomAttributes(typeof(TestMethodAttribute)).Any();
        }

        public string Name => UnitTesType.Name;

        private object _instance;
        public Type UnitTesType { get;  }

        public IEnumerable<MethodInfo> TestMethods { get; }

        public object Instance {
            get {
                if (_instance == null) {
                    _instance = Activator.CreateInstance(UnitTesType);
                }
                return _instance;
            }
        }

        public void Test() {
            var instance = Activator.CreateInstance(UnitTesType);
            foreach (var testMethod in TestMethods) {
                
            }
        }
    }
}