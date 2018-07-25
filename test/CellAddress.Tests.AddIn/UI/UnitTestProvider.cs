using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CellAddressTests.AddIn.UI {
    public class UnitTestProvider {
        

        public UnitTestProvider() {

        }

        public IEnumerable<TestCategory> TestCategories { get; private set; }

        public void Resolve() {
            // AppDomain.CurrentDomain.GetAssemblies().Where(a=>!a.IsDynamic)
            var unitTypes = GetUnitTestTypes(typeof(UnitTestProvider).Assembly);
            TestCategories = unitTypes.Select(t => new TestCategory(t));
        }

        private static IEnumerable<Type> GetUnitTestTypes(Assembly assembly) {
            return  assembly.GetExportedTypes().Where(IsUnitTestClass).ToArray();
        }

        private static IEnumerable<MethodInfo> GetMethods(Type testType) {
            var methods = typeof(CellAddRressExtensionsTests).GetMethods(BindingFlags.Instance | BindingFlags.Public);
            return methods.Where(IsUnitTestMethod);
        }

        private static bool IsUnitTestMethod(MethodInfo method) {
            return method.GetParameters().Length == 0 && method.GetCustomAttributes(typeof(TestMethodAttribute)).Any();
        }

        private static bool IsUnitTestClass(Type type) {
            return type.GetCustomAttributes<TestClassAttribute>().Any();
        }
    }
}