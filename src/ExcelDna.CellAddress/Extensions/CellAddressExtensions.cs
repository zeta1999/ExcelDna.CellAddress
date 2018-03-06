using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna.Extensions {
    public static class CellAddressExtensions {
        /// <summary>
        /// 根据索引返回第n个单元格
        /// </summary>
        /// <param name="range"></param>
        /// <param name="index"></param>
        /// <param name="direction">排列顺序(行优先/列优先)</param>
        /// <returns></returns>
        public static CellAddress GetCell(this CellAddress range, int index, XlFillDirection direction = XlFillDirection.RowFirst) {
            if (index >= range.Count && index < -1) {
                throw new IndexOutOfRangeException($"索引超出范围,-1< index < {range.Count}");
            }
            if (range.Count == 1 && index == 0) {
                return range;
            }
            if (direction == XlFillDirection.ColumnFirst) {
                //列优先
                return range.GetCell(index % range.Rows, index / range.Rows);
            }
            return range.GetCell(index % range.Columns, index / range.Columns);
        }

        public static CellAddress GetCell(this CellAddress cell, int rowIndex = 0, int columnIndex = 0) {
            return new CellAddress(cell.SheetName, cell.RowFirst + rowIndex, cell.ColumnFirst + columnIndex);
        }

        /// <summary>
        /// 返回 <see cref="CellAddress"/> 对象，它代表位于指定单元格区域的一定的偏移量位置上的区域。
        /// 返回的区域和原始区域大小相同
        /// <remarks>
        /// <seealso cref="Microsoft.Office.Interop.Excel.Range.Offset"/>
        /// </remarks>
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowOffset">
        /// 区域偏移的行数（正数、负数或 0（零））。正数表示向下偏移，负数表示向上偏移。默认值是 0
        /// </param>
        /// <param name="columnOffset">
        /// 
        /// </param>
        /// <returns></returns>
        public static CellAddress Offset(this CellAddress cell, int rowOffset = 0, int columnOffset = 0) {
            return new CellAddress(cell.SheetName,
                cell.RowFirst + rowOffset,
                cell.RowLast + rowOffset,
                cell.ColumnFirst + columnOffset,
                cell.ColumnLast + columnOffset);
        }

        /// <summary>
        /// 单元格地址最大的一个
        /// </summary>
        /// <param name="cell1"></param>
        /// <param name="cell2"></param>
        /// <returns></returns>
        public static CellAddress Max(this CellAddress cell1, CellAddress cell2) {
            if (cell1.ColumnFirst > cell2.ColumnFirst || cell1.RowFirst > cell2.RowFirst) {
                return cell1;
            } else {
                return cell2;
            }
        }

        /// <summary>
        /// 单元格地址 最小的一个
        /// </summary>
        /// <param name="cell1"></param>
        /// <param name="cell2"></param>
        /// <returns></returns>
        public static CellAddress Min(this CellAddress cell1, CellAddress cell2) {
            if (cell1.ColumnFirst > cell2.ColumnFirst || cell1.RowFirst > cell2.RowFirst) {
                return cell2;
            } else {
                return cell1;
            }
        }


        #region CellAddress Values
        /// <summary>
        /// 从单元格读取数据
        /// </summary>
        /// <returns></returns>
        public static T GetValue<T>(this CellAddress address) {
            var reference = address.CellReference;
            if (reference.IsEmpty()) {
                return default(T);
            }
            return reference.GetValue<T>();
        }
        /// <summary>
        /// 从单元格读取数据
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<T> GetValues<T>(this CellAddress address) {
            if (address.Count == 1) {
                return new T[] { address.GetValue<T>() };
            }
            if (address.HasRange) {
                var values = address.CellRange.Value2 as object[,];
                return values.AsIEnumerable<T>();
            }
            var reference = address.CellReference;
            if (!reference.IsEmpty()) {
                return reference.GetValues<T>();
            }
            return new T[0];
        }

        /// <summary>
        /// 设置单元格内容
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void SetValue(this CellAddress cell, object value) {
            if (value.IsNull()) {
                cell.SetValueInternal(new object[cell.Rows, cell.Columns]);
            } else if (cell.Count == 1) {
                var vt = new object[1, 1];
                vt[0, 0] = value;
                cell.SetValueInternal(vt);
            } else {
                if (value is object[,] array) {
                    cell.SetValueInternal(array);
                } else if (value is string) {
                    var str = (string)value;
                    var arr = str.Split(',');
                    var vt = arr.ToMatrix(cell.Rows, cell.Columns);
                    cell.SetValueInternal(vt);
                }
            }
        }

        private static void SetValueInternal(this CellAddress cell, object[,] value) {
            var reference = cell.CellReference;
            try {
                if (reference != null) {
                    reference.SetValue(value);
                } else {
                    cell.CellRange.Value2 = value;
                }
            } catch (XlCallException) {
                throw new Exception($"{cell}单元格定义错误,无法写入该单元格");
            }
        }

        /// <summary>
        /// 设置公式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="formula"></param>
        public static void SetFormula(this CellAddress cell, string formula) {
            cell.CellReference.SetFormula(formula);
        }

        public static void ClearContents(this CellAddress cell) {
            if (cell == null) {
                throw new ArgumentNullException(nameof(cell));
            }
            cell.CellRange.ClearContents();
        }

        #endregion

        public static void Activate(this CellAddress cell) {
            cell?.CellReference.Activate();
        }

        internal static Range GetRange(this string celladdress) {
            try {
                var xlApp = ExcelDnaUtil.Application;
                if (!(xlApp is Application application)) {
                    throw new NullReferenceException();
                }
                return application.Range[celladdress];
            } catch (InvalidOperationException ioe) {
                //当前 ExcelApplication 不可用
                Trace.TraceWarning("GetRange Error {0}", ioe);
                throw ioe;
            }
        }

        internal static CellAddress GetRange(this IEnumerable<CellAddress> cells) {
            if (cells == null) {
                return CellAddress.Ref;
            }

            var cellArray = cells as CellAddress[] ?? cells.ToArray();
            if (!cellArray.Any()) {
                return CellAddress.Ref;
            }
            if (cellArray.Length == 1) {
                return cellArray[0];
            }

            var sheet = cellArray.Select(c => c.SheetName).FirstOrDefault();
            var rowFirst = cellArray.Min(c => c.RowFirst);
            var rowLast = cellArray.Max(c => c.RowLast);
            var colFirst = cellArray.Min(c => c.ColumnFirst);
            var colLast = cellArray.Max(c => c.ColumnLast);

            return new CellAddress(sheet, rowFirst, rowLast, colFirst, colLast);
        }
    }
}