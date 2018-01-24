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
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;


namespace ExcelDna.Extensions{
    /// <summary>
    /// <see cref="ExcelReference"/> 扩展方法
    /// </summary>
    public static class ExcelReferenceExtensions{

        #region ExcelReference 基本扩展方法

        public static T GetValue<T>(this ExcelReference range){
            return range.GetValue().ConvertTo<T>();
        }

        public static IEnumerable<T> GetValues<T>(this ExcelReference range)
        {
            return range.Cells().Select(c => c.GetValue()).OfType<T>();
        }

        public static IEnumerable<T> GetValues<T>(this IEnumerable<ExcelReference> ranges)
        {
            return ranges.Cells().Select(c => c.GetValue()).OfType<T>();
        }

        private static bool IsSingleCell(this ExcelReference reference) {
            return reference.ColumnFirst == reference.ColumnLast && reference.RowFirst == reference.RowLast;
        }

        /// <summary>
        ///     单元格范围地址
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string Address(this ExcelReference range){
            string sheetName = range.SheetNameLocal();

            if (range.IsSingleCell()) {
                return $"{sheetName}!${range.ColumnFirst.ColumnName()}${range.RowFirst + 1}";
            }
            return
                $"{sheetName}!${range.ColumnFirst.ColumnName()}${range.RowFirst + 1}:${range.ColumnLast.ColumnName()}${range.RowLast + 1}";

            //return string.Join(",", range.InnerReferences.Select(r => r.Address()));
        }

        /// <summary>
        ///     ExcelReference 本地地址 (不包括 Worksheet 名称)
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string AddressLocal(this ExcelReference range){
            if (range.IsSingleCell()){
                    return $"${range.ColumnFirst.ColumnName()}${range.RowFirst + 1}";
               }
                return
                    $"${range.ColumnFirst.ColumnName()}${range.RowFirst + 1}:${range.ColumnLast.ColumnName()}${range.RowLast + 1}";
        }

        /// <summary>
        ///     根据 列索引 计算列名 0=>A  1=>B
        ///     从 0 开始计算
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        private static string ColumnName(this int columnIndex){
            if (columnIndex < 0){
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "列索引不能小于 0");
            }
            if (columnIndex < 26){
                return ((char) ('A' + columnIndex)).ToString();
            }
            if (columnIndex < 26 + (26*26)){
                var nameChars = new char[2];
                nameChars[0] = (char) ('A' + (int) (Math.Floor((double) columnIndex/26)) - 1);
                nameChars[1] = (char) ('A' + (columnIndex%26));
                return new string(nameChars);
            }
            if (columnIndex < 26 + (26*26) + (26*26*26)){
                double modValue = columnIndex - 26;
                var nameChars = new char[3];
                nameChars[0] = (char) ('A' + (int) (Math.Floor(modValue/(26*26))) - 1);
                modValue = modValue%(26*26);
                nameChars[1] = (char) ('A' + (int) (Math.Floor(modValue/26)));
                nameChars[2] = (char) ('A' + (columnIndex%26));
                return new string(nameChars);
            }
            throw new ArgumentException("超出列名范围");
        }


        /// <summary>
        /// 激活单元格
        /// </summary>
        /// <param name="reference"></param>
        public static void Activate(this ExcelReference reference) {
            XlCall.Excel(XlCall.xlcSelect, reference);
        }

        /// <summary>
        /// 选中单元格
        /// </summary>
        /// <param name="range"></param>
        public static void Select(this ExcelReference range){
            XlCall.Excel(XlCall.xlcFormulaGoto, range);
            XlCall.Excel(XlCall.xlcSelect, range, Type.Missing);
        }

        public static void Copy(this ExcelReference fromRange, ExcelReference toRange = null){
            object to_range = (toRange == null) ? Type.Missing : toRange;
            XlCall.Excel(XlCall.xlcCopy, fromRange, to_range);
        }

        public static ExcelReference Resize(this ExcelReference range, int rows, int cols){
            rows = (rows < 1) ? 1 : rows;
            cols = (cols < 1) ? 1 : cols;
            return new ExcelReference(range.RowFirst, range.RowFirst + rows - 1, range.ColumnFirst,
                range.ColumnFirst + cols - 1, range.SheetId);
        }

        /// <summary>
        ///     从当前单元格计算偏移范围，返回和原始单元格相同大小的 ExcelReference 对象
        /// </summary>
        /// <param name="range"></param>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        public static ExcelReference Offset(this ExcelReference range, int rows, int cols){
            return new ExcelReference(range.RowFirst + rows, range.RowLast + rows, range.ColumnFirst + cols,
                range.ColumnLast + cols, range.SheetId);
        }

        /// <summary>
        ///     返回从当前单元格 左上角开始计算的 范围，根据给定的行列值，返回一个 rows*cols 大小的单元格区间
        /// </summary>
        /// <param name="range"></param>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        public static ExcelReference Range(this ExcelReference range, int rows, int cols){
            if (rows == 0){
                rows = 1;
            } else if (rows > 0){
                rows--;
            }
            if (cols == 0){
                cols = 1;
            } else if (cols > 0){
                cols--;
            }
            return new ExcelReference(range.RowFirst, range.RowFirst + rows, range.ColumnFirst,
                range.ColumnFirst + cols, range.SheetId);
        }

        /// <summary>
        ///     根据给定参考范围，返回给定单元格相关的地址范围
        ///     如果 <see cref="address">单元格地址</see>包括工作簿，则使用地址的工作簿
        ///     如果地址中没有指定工作簿，则使用 <see cref="refRange">参考单元格</see>的工作簿
        /// </summary>
        /// <param name="refRange"></param>
        /// <param name="address"></param>
        /// <returns></returns>
        public static IEnumerable<ExcelReference> Range(this ExcelReference refRange, string address){
            if (string.IsNullOrEmpty(address)){
                yield return null;
            } else{
                if (address.StartsWith("=")){
                    address = address.Substring(address.IndexOf('=') + 1);
                }
                string[] addressArray;
                if (address.IndexOf(',') > 0){
                    //多个单元格范围
                    addressArray = address.Split(',');
                } else{
                    addressArray = new[]{address};
                }
                string refSheet = refRange.SheetNameLocal();
                foreach (string item in addressArray){
                    if (item.IndexOf('!') > 0){
                        //包含工作簿定义
                        yield return XlCall.Excel(XlCall.xlfEvaluate, item) as ExcelReference;
                    } else{
                        yield return
                            XlCall.Excel(XlCall.xlfEvaluate, string.Format("{0}!{1}", refSheet, item)) as ExcelReference;
                    }
                }
            }
        }

        public static int CellsCount(this ExcelReference range) {
            return range.Rows()*range.Columns();
        }

        /// <summary>
        ///     返回 范围内的 单元格集合，按照 列 优先
        ///     该方法 不支持 合并单元格检测
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static IEnumerable<ExcelReference> Cells(this ExcelReference range){
            if (range.IsSingleCell()){
                return new[]{range};
            }
            return range.InnerReferences.Cells();
        }

        /// <summary>
        ///     返回 范围内的 单元格集合，按照 列 优先
        ///     该方法 不支持 合并单元格检测
        /// </summary>
        /// <param name="ranges"></param>
        /// <returns></returns>
        public static IEnumerable<ExcelReference> Cells(this IEnumerable<ExcelReference> ranges){
            ExcelReference[] references = ranges as ExcelReference[] ?? ranges.ToArray();
            if (!references.Any()){
                yield return null;
            }
            if (references.Count() == 1){
                ExcelReference range = references.FirstOrDefault();
                int columns = range.Columns();
                int rows = range.Rows();
                for (int c = 0; c < columns; c++){
                    for (int r = 0; r < rows; r++) {
                        var row = range.RowFirst + r;
                        var col = range.ColumnFirst + c;
                        yield return new ExcelReference(row,row,col,col,range.SheetId);
                    }
                }
            } else{
                int cellsCount = references.Sum(r => r.CellsCount());
                int i = 0;
                int col = 0;
                while (i < cellsCount){
                    foreach (ExcelReference range in references){
                        int columns = range.Columns();
                        int rows = range.Rows();
                        if (col < columns){
                            for (int row = 0; row < rows; row++){
                                yield return
                                    new ExcelReference(range.RowFirst + row, range.RowFirst + row, range.ColumnFirst + col, range.ColumnFirst + col, range.SheetId);
                            }
                        }
                    }
                    col++;
                    i++;
                }
            }
        }

        /// <summary>
        ///     单元格是否为空
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool IsEmpty(this ExcelReference range){
            if (range == null){
                return true;
            }
            object value = range.GetValue();
            if (value is ExcelEmpty || value is ExcelError || value is ExcelMissing){
                return true;
            }
            return string.IsNullOrEmpty(value.ToString());
        }

        /// <summary>
        ///     向单元格赋值,横向 单行，多列 [n,0]
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="range"></param>
        /// <param name="values"></param>
        public static void SetValues<T>(this ExcelReference range, IEnumerable<T> values){
            if (range.CellsCount() > 1){
                range.SetValue(values.ToArray());
            } else{
                T[] array = values.ToArray();
                int count = array.Length;
                var colRange = new ExcelReference(range.RowFirst, range.RowFirst,
                    range.ColumnFirst, range.ColumnLast + count - 1
                    );
                colRange.SetValue(array);
            }
        }

        /// <summary>
        ///     向单元格赋值,横向 单行，多列 [n,0]
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="range"></param>
        /// <param name="values"></param>
        public static void SetValues<T>(this ExcelReference range, object[] values){
            if (range.Columns() > 1){
                range.SetValue(values.ToArray());
            } else{
                int count = values.Length;
                var colRange = new ExcelReference(range.RowFirst, range.RowFirst,
                    range.ColumnFirst, range.ColumnLast + count - 1
                    );
                colRange.SetValue(values);
            }
        }

        /// <summary>
        ///     描述 <see cref="cell" /> 单元格是否包含在<see cref="range" /> 范围内
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="range">可能包含 cell 的单元格范围</param>
        /// <returns></returns>
        public static bool InRange(this ExcelReference cell, ExcelReference range){
            if (range == null || cell == null){
                return false;
            }
            if (range.SheetId != cell.SheetId){
                return false;
            }
            return range.RowLast >= cell.RowLast
                   && range.RowFirst <= cell.RowFirst
                   && range.ColumnLast >= cell.ColumnLast
                   && range.ColumnFirst <= cell.ColumnFirst;
        }

        #region 基本属性

        /// <summary>
        ///     单元格范围的 列数
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int Columns(this ExcelReference range){
            return range.ColumnLast - range.ColumnFirst + 1;
        }

        /// <summary>
        ///     单元格范围 的行数
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static int Rows(this ExcelReference range){
            return range.RowLast - range.RowFirst + 1;
        }

        #endregion 基本属性

        #region Formula

        /// <summary>
        ///     获取 单元格公式
        ///     <see cref="XlCall.xlfGetCell" />
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        /// <remarks>
        ///     6	Formula in reference, as text, in either A1 or R1C1 style depending on the workspace setting.
        /// </remarks>
        public static string GetFormula(this ExcelReference range){
            try{
                return (string) XlCall.Excel(XlCall.xlfGetCell, 6, range);
            } catch (XlCallException){
                //给定单元格没有公式
                return string.Empty;
            }
        }

        /// <summary>
        ///     清除单元格公式
        /// </summary>
        /// <param name="range"></param>
        public static void ClearFormula(this ExcelReference range){
            IEnumerable<ExcelReference> cells = range.Cells();
            foreach (ExcelReference cell in cells){
                //删除公式
                XlCall.Excel(XlCall.xlcFormula, string.Empty, cell);
            }
        }

        /// <summary>
        ///     设置单元格公式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="formula"></param>
        public static void SetFormula(this ExcelReference range, string formula){
            if (string.IsNullOrEmpty(formula)){
                //删除公式
                range.ClearFormula();
                return;
            }
            // Get the formula and convert to R1C1 mode
            var isR1C1Mode = (bool) XlCall.Excel(XlCall.xlfGetWorkspace, 4);
            string formulaR1C1 = formula;
            if (!isR1C1Mode){
                formulaR1C1 = (string)
                    XlCall.Excel(XlCall.xlfFormulaConvert, formula, true, false, ExcelMissing.Value, range);
            }
            object ignoredResult;
            XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlcFormula, out ignoredResult, formulaR1C1, range);
            if (retval != XlCall.XlReturn.XlReturnSuccess){
                // TODO: Consider what to do now!?
                // Might have failed due to array in the way.
                range.SetValue("'" + formula);
            }
        }

        /// <summary>
        ///     获取指定单元格是否包含公式
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool HasFormula(this ExcelReference range){
            var formula = (string) XlCall.Excel(XlCall.xlfGetCell, 41, range);
            return !string.IsNullOrEmpty(formula);
        }

        #endregion Formula

        #region Next and NextRows

        /// <summary>
        ///     右侧下一个单元格 不包括当前单元格
        ///     遇到空单元格结束
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static IEnumerable<ExcelReference> Next(this ExcelReference range){
            return range.Next(1);
        }

        /// <summary>
        ///     右侧下一个单元格 不包括当前单元格
        ///     遇到空单元格结束
        /// </summary>
        /// <param name="range"></param>
        /// <param name="skip">起始相对位置,如果为0 则从当前单元格开始</param>
        /// <returns></returns>
        public static IEnumerable<ExcelReference> Next(this ExcelReference range, int skip){
            int i = skip;
            while (true){
                var nextCell =
                    new ExcelReference(range.RowLast,range.RowLast, range.ColumnLast + i,range.ColumnLast + i, range.SheetId);
                if (nextCell.IsEmpty()){
                    break;
                }
                yield return nextCell;
                i++;
            }
        }

        /// <summary>
        ///     下方 下一个单元格 不包括当前单元格
        ///     遇到空单元格结束
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static IEnumerable<ExcelReference> NextRows(this ExcelReference range){
            return range.NextRows(1);
        }

        /// <summary>
        ///     下方 下一个单元格
        ///     遇到空单元格结束
        /// </summary>
        /// <param name="range"></param>
        /// <param name="skip">起始相对位置,如果为0 则从当前单元格开始</param>
        /// <returns></returns>
        public static IEnumerable<ExcelReference> NextRows(this ExcelReference range, int skip){
            int i = skip;
            while (true){
                var nextCell =
                    new ExcelReference(range.RowLast + i, range.RowLast + i, range.ColumnLast, range.ColumnLast, range.SheetId);
                if (nextCell.IsEmpty()){
                    break;
                }
                yield return nextCell;
                i++;
            }
        }

        #endregion Next and NextRows

        #region Worksheet

        /// <summary>
        ///     获取 单元格所在 工作表的 名称
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string SheetRef(this ExcelReference range){
            return (string) XlCall.Excel(XlCall.xlfGetCell, 62, range);
        }

        /// <summary>
        ///     单元格所在 工作表名称,包括 Workbook 名称
        ///     [BookName]SheetName
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string SheetName(this ExcelReference range){
            return (string) XlCall.Excel(XlCall.xlSheetNm, range);
        }

        public static string WorkbookName(this ExcelReference range){
            return (string) XlCall.Excel(XlCall.xlfGetCell, 66, range);
        }

        /// <summary>
        ///     单元格所在 工作表 本地名称,不包括 Workbook 名称
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string SheetNameLocal(this ExcelReference range){
            string sheetName = range.SheetName();
            return sheetName.Substring(sheetName.IndexOf(']') + 1);
        }

        /// <summary>
        ///     给定单元格所在 工作簿中 是否包含 该工作表
        /// </summary>
        /// <param name="range"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static bool HasWorksheet(this ExcelReference range, string sheetName){
            string workName = range.WorkbookName();
            try{
                object result = XlCall.Excel(XlCall.xlfGetDocument, 76, string.Format("[{0}]{1}", workName, sheetName));
                return !result.IsNull();
            } catch (Exception){
                return false;
            }
        }

        #endregion Worksheet

        #endregion ExcelReference 扩展方法
    }
}