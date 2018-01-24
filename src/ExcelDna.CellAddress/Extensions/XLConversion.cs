/*
The MIT License (MIT)

Copyright (c) 2014 Joachim Loebb

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using RedRiver.ExcelDNA.Extensions;

namespace ExcelDna.Extensions{
    public static class XlConversion{
        #region object conversion

        //a thread-safe way to hold default instances created at run-time
        private static readonly ConcurrentDictionary<Type, object> typeDefaults =
            new ConcurrentDictionary<Type, object>();

        private static object GetDefault(Type type){
            return type.IsValueType ? typeDefaults.GetOrAdd(type, t => Activator.CreateInstance(t)) : null;
        }

        public static object ConvertTo(this object vt, Type toType){
            Type fromType = vt.GetType();

            if (vt == null){
                return GetDefault(toType);
            }
            if (fromType == typeof (DBNull)){
                return GetDefault(toType);
            }

            if (fromType == typeof (ExcelEmpty) || fromType == typeof (ExcelError) || fromType == typeof (ExcelMissing)){
                return GetDefault(toType);
            }

            if (fromType == typeof (ExcelReference)){
                var r = (ExcelReference) vt;
                object val = r.GetValue();
                return ConvertTo(val, toType);
            }

            //acount for nullable types
            toType = Nullable.GetUnderlyingType(toType) ?? toType;

            if (toType == typeof (DateTime)){
                DateTime dt = DateTime.FromOADate(0.0);
                if (fromType == typeof (DateTime)){
                    dt = (DateTime) vt;
                } else if (fromType == typeof (double)){
                    dt = DateTime.FromOADate((double) vt);
                } else if (fromType == typeof (string)){
                    DateTime result;
                    if (DateTime.TryParse((string) vt, out result)){
                        dt = result;
                    }
                }
                return Convert.ChangeType(dt, toType);
            }
            if (toType == typeof (XlDate)){
                XlDate dt = 0.0;
                if (fromType == typeof (DateTime)){
                    dt = (DateTime) vt;
                } else if (fromType == typeof (double)){
                    dt = (double) vt;
                } else if (fromType == typeof (string)){
                    DateTime result;
                    if (DateTime.TryParse((string) vt, out result)){
                        dt = result;
                    } else{
                        dt = 0.0;
                    }
                } else{
                    dt = (double) Convert.ChangeType(vt, typeof (double));
                }
                return Convert.ChangeType(dt, toType);
            }
            if (toType == typeof (double)){
                double dt = 0.0;
                if (fromType == typeof (double)){
                    dt = (double) vt;
                } else if (fromType == typeof (DateTime)){
                    dt = ((DateTime) vt).ToOADate();
                } else if (fromType == typeof (string)){
                    double.TryParse((string) vt, out dt);
                } else{
                    dt = (double) Convert.ChangeType(vt, typeof (double));
                }
                return Convert.ChangeType(dt, toType);
            }
            if (toType.IsEnum){
                try{
                    return Enum.Parse(toType, vt.ToString(), true);
                } catch (Exception){
                    return GetDefault(toType);
                }
            }
            return Convert.ChangeType(vt, toType);
        }

        /// <summary>
        /// 类型是否为空
        /// Null 包括 <see cref="DBNull"/>,<see cref="ExcelEmpty"/><see cref="ExcelError"/> <see cref="ExcelMissing"/>
        /// <see cref="Type.Missing"/>
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        public static bool IsNull(this object instance) {
            return instance == null || instance == ExcelMissing.Value || instance == ExcelMissing.Value
                   || instance == System.Type.Missing
                   || instance is ExcelEmpty || instance is ExcelError || instance is ExcelMissing || instance is DBNull;
        }

        public static bool IsNullOrEmpty(this object instance){
            if (instance.IsNull()){
                return true;
            }
            var array = instance as Array;
            if (array != null && array.Length == 0){
                return true;
            }
            var str = instance as string;
            if (str != null){
                return string.IsNullOrEmpty(str);
            }
            return false;
        }

        /// <summary>
        /// 类型转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="vt"></param>
        /// <returns></returns>
        public static T ConvertTo<T>(this object vt){
            if (vt.IsNull()){
                return default(T);
            }

            Type toType = typeof (T);
            Type fromType = vt.GetType();
            
            var reference = vt as ExcelReference;
            if (reference != null){
                object val = reference.GetValue();
                return ConvertTo<T>(val);
            }

            //acount for nullable types
            toType = Nullable.GetUnderlyingType(toType) ?? toType;

            if (toType == typeof (DateTime)){
                DateTime dt = DateTime.FromOADate(0.0);
                if (fromType == typeof (DateTime)){
                    dt = (DateTime) vt;
                } else if (fromType == typeof (double)){
                    dt = DateTime.FromOADate((double) vt);
                } else if (fromType == typeof (string)){
                    DateTime result;
                    if (DateTime.TryParse((string) vt, out result)){
                        dt = result;
                    }
                }
                //note this will work also if T is nullable
                return (T) Convert.ChangeType(dt, toType);
            }
            if (toType == typeof (XlDate)){
                XlDate dt = 0.0;
                if (fromType == typeof (DateTime)){
                    dt = (DateTime) vt;
                } else if (fromType == typeof (double)){
                    dt = (double) vt;
                } else if (fromType == typeof (string)){
                    DateTime result;
                    if (DateTime.TryParse((string) vt, out result)){
                        dt = result;
                    } else{
                        dt = 0.0;
                    }
                } else{
                    dt = (double) Convert.ChangeType(vt, typeof (double));
                }
                return (T) Convert.ChangeType(dt, toType);
            }
            if (toType == typeof (double)){
                double dt = 0.0;
                if (fromType == typeof (double)){
                    dt = (double) vt;
                } else if (fromType == typeof (DateTime)){
                    dt = ((DateTime) vt).ToOADate();
                } else if (fromType == typeof (string)){
                    double.TryParse((string) vt, out dt);
                } else{
                    dt = (double) Convert.ChangeType(vt, typeof (double));
                }
                return (T) Convert.ChangeType(dt, toType);
            }
            if (toType.IsEnum){
                try{
                    return (T) Enum.Parse(typeof (T), vt.ToString(), true);
                } catch (Exception){
                    return default(T);
                }
            }
            return (T) Convert.ChangeType(vt, toType);
        }

        public static void ConvertVT<T>(this object vt, out T value){
            value = vt.ConvertTo<T>();
        }

        public static T[] ToVector<T>(this object vt){
            if (vt is Array){
                return ToVector<T>(vt as object[,]);
            }

            var retval = new T[1];
            vt.ConvertVT(out retval[0]);

            return retval;
        }

        public static T[] ToVector<T>(this object[,] vt){
            int n = vt.GetLength(0), k = vt.GetLength(1);
            int l = 0;

            var @out = new T[n*k];

            for (int i = 0; i < n; i++){
                for (int j = 0; j < k; j++){
                    vt.GetValue(i, j).ConvertVT(out @out[l]);
                    l++;
                }
            }

            return @out;
        }

        public static object[] ToArray(this object[,] vt){
            return vt.ToIEnumerable().ToArray();
        }

        /// <summary>
        /// 二维数组转换为 枚举集合
        /// </summary>
        /// <param name="vt"></param>
        /// <returns></returns>
        public static IEnumerable<T> AsIEnumerable<T>(this object[,] vt) {
            foreach (var item in vt) {
                yield return item.ConvertTo<T>();
            }
        }


        private static IEnumerable<object> ToIEnumerable(this object[,] vt) {
            foreach (var item in vt){
                yield return item;
            }
        }

        /// <summary>
        /// 把一个枚举对象转换为一个二维数组
        /// 排列顺序为先行后列
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="values"></param>
        /// <param name="rows">变换数组的行数，rows 不能小于1</param>
        /// <param name="columns">变换数组的列数，columns 不能小于 1</param>
        /// <param name="emptyValue"></param>
        /// <returns>变换后的二维数组 T[rows,columns]</returns>
        public static T[,] ToMatrix<T>(this IEnumerable<T> values, int rows, int columns, T emptyValue = default(T)) {
            if (values == null) {
                //throw new ArgumentNullException(nameof(values));
                values = new T[0];
            }
            if (rows < 1) {
                throw new ArgumentOutOfRangeException(nameof(rows));
            }
            if (columns < 1) {
                throw new ArgumentOutOfRangeException(nameof(columns));
            }
            var matrix = new T[rows, columns];
            using (var enumerator = values.GetEnumerator()) {
                for (int r = 0; r < rows; r++) {
                    for (int c = 0; c < columns; c++) {
                        if (enumerator.MoveNext()) {
                            if (enumerator.Current.IsNull()) {
                                matrix[r, c] = emptyValue;
                            } else {
                                matrix[r, c] = enumerator.Current;
                            }
                        } else {
                            matrix[r, c] = emptyValue;
                        }
                    }
                }
            }
            return matrix;
        }

        /// <summary>
        /// 把一个枚举对象转换为一个二维数组
        /// 排列顺序为先列后行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="values"></param>
        /// <param name="rows">变换数组的行数，rows 不能小于1</param>
        /// <param name="columns">变换数组的列数，columns 不能小于 1</param>
        /// <param name="emptyValue">默认空值</param>
        /// <returns>变换后的二维数组 T[rows,columns]</returns>
        public static T[,] ToMatrixA<T>(this IEnumerable<T> values, int rows, int columns, T emptyValue = default(T)) {
            if (values == null) {
                throw new ArgumentNullException(nameof(values));
            }
            if (rows < 1) {
                throw new ArgumentOutOfRangeException(nameof(rows));
            }
            if (columns < 1) {
                throw new ArgumentOutOfRangeException(nameof(columns));
            }
            var matrix = new T[rows, columns];

            using (var enumerator = values.GetEnumerator()) {
                for (int c = 0; c < columns; c++) {
                    for (int r = 0; r < rows; r++) {
                        if (enumerator.MoveNext()) {
                            if (enumerator.Current.IsNull()) {
                                matrix[r, c] = emptyValue;
                            } else {
                                matrix[r, c] = enumerator.Current;
                            }
                        } else {
                            matrix[r, c] = emptyValue;
                        }
                    }
                }
            }
            return matrix;
        }


        public static T[,] ToMatrix<T>(this object vt){
            if (vt is Array){
                return ToMatrix<T>(vt as object[,]);
            }

            var retval = new T[1, 1];
            vt.ConvertVT(out retval[0, 0]);

            return retval;
        }

        public static T[,] ToMatrix<T>(this object[,] vt){
            int n = vt.GetLength(0), k = vt.GetLength(1);

            var @out = new T[n, k];

            for (int i = 0; i < n; i++){
                for (int j = 0; j < k; j++){
                    vt.GetValue(i, j).ConvertVT(out @out[i, j]);
                }
            }

            return @out;
        }

        public static object ToVariant<T>(this T[,] vt){
            int n = vt.GetLength(0), k = vt.GetLength(1);
            var @out = new object[n, k];

            for (int i = 0; i < n; i++){
                for (int j = 0; j < k; j++){
                    @out[i, j] = vt.GetValue(i, j);
                }
            }

            return @out;
        }

        #endregion
    }
}