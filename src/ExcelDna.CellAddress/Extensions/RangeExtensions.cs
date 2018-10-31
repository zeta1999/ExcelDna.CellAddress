using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna.Extensions {
    /// <summary>
    ///     Excel Range 扩展方法
    /// </summary>
    public static class RangeExtensions {
/*
        /// <summary>
        ///     <see cref="Range">Range</see>对象转换为 单一<see cref="ExcelReference" />
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static ExcelReference ToReference(this Range cell) {
            var row = cell.Row - 1;
            var col = cell.Column - 1;
            return new ExcelReference(row, row, col, col, cell.Worksheet.Name);
        }
*/

/*
        /// <summary>
        ///    判断单元格是否为空
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool IsEmpty(this Range range) {
            if (range == null) {
                return true;
            }

            if (range.Count > 1) {
                //多个单元格
                var values = range.Value2 as object[,];
                foreach (var item in values) {
                    if (item.IsEmpty()) {
                        return true;
                    }
                }
                return false;
            }
            object value = range.Value2;
            return value.IsEmpty();
        }
*/

/*
        private static bool IsEmpty(this object value) {
            if (value == null || value is DBNull || value is ExcelEmpty || value is ExcelError ||
                value is ExcelMissing
                || value == Type.Missing) {
                return true;
            }
            return string.IsNullOrEmpty(value.ToString());
        }
*/

/*
        /// <summary>
        ///     获取下一行单元格,包括合并单元格
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static Range GetNextRow(this Range cell) {
            if ((bool)cell.MergeCells) {
                return cell.Next[cell.MergeArea.Rows.Count + 1, 0] as Range;
            }
            return cell.Next[2, 0] as Range;
        }
*/


        #region AsIEnumerable

/*
        public static IEnumerable<Range> AsEnumerable(this Areas areas) {
            foreach (Range area in areas) {
                yield return area;
            }
        } 
*/

        #endregion

        #region GetCells

/*
        /// <summary>
        ///     返回 <see cref="Microsoft.Office.Interop.Excel.Range" /> 对象中单元格的集合
        ///     按照列优先的顺序返回，支持 合并单元格检测
        /// </summary>
        /// <param name="areas"></param>
        /// <returns></returns>
        public static IEnumerable<Range> GetCells(this Areas areas) {
            if (areas == null) {
                throw new ArgumentNullException(nameof(areas));
            }
            foreach (Range area in areas) {
                foreach (Range column in area.Columns) {
                    foreach (Range cell in column.Rows) {
                        if ((bool)cell.MergeCells) {
                            if (cell.Address == cell.MergeArea.Offset[0, 0].Address) {
                                yield return cell;
                            }
                        } else {
                            yield return cell;
                        }
                    }
                }
            }
        }
*/

/*
        /// <summary>
        ///     返回 <see cref="Microsoft.Office.Interop.Excel.Range" /> 对象中单元格的集合
        ///     按照列优先的顺序返回，支持 合并单元格检测
        /// </summary>
        /// <param name="range"></param>
        /// <param name="direction">遍历方向,默认为 列优先</param>
        /// <returns></returns>
        public static IEnumerable<Range> GetCells(this Range range, XlFillDirection direction = XlFillDirection.ColumnFirst ) {
            if (range == null) {
                throw new ArgumentNullException(nameof(range));
            }
            if (direction == XlFillDirection.ColumnFirst) {
                foreach (Range area in range.Areas) {
                    foreach (Range column in area.Columns) {
                        foreach (Range cell in column.Rows) {
                            if ((bool)cell.MergeCells) {
                                if (cell.Address == cell.MergeArea.Offset[0].Address) {
                                    yield return cell.MergeArea;
                                }
                            } else {
                                yield return cell;
                            }
                        }
                    }
                }
            } else {
                foreach (Range cell in range.Cells) {
                    if ((bool)cell.MergeCells) {
                        if (cell.Address == cell.MergeArea.Offset[0].Address) {
                            yield return cell.MergeArea;
                        }
                    } else {
                        yield return cell;
                    }
                }
            }
        }
*/

/*
        /// <summary>
        ///     列区域，遍历各个区域的 每个列
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private static IEnumerable<Range> ColumnAreas(this Range range) {
            var areaCount = range.Areas.Count;
            var areas = 0;
            var column = 0;
            while (true) {
                foreach (Range area in range.Areas) {
                    if (column < area.Columns.Count) {
                        yield return area.Columns[column + 1] as Range;
                    } else {
                        areas++;
                    }
                }
                column++;
                if (areas >= areaCount) {
                    break;
                }
            }
        }
*/

        #endregion GetCells

        #region GetFormula

/*
        /// <summary>
        ///     设置 单元格公式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="formula"></param>
        /// <returns></returns>
        public static void SetFormula(this Range range, string formula) {
            if (string.IsNullOrEmpty(formula)) {
                return;
            }
            if (!formula.StartsWith("=")) {
                formula = "=" + formula;
            }
            range.ClearContents();
            range.Formula = formula;
            range.FormulaHidden = true;
        }
*/

/*
        /// <summary>
        ///     设置 单元格公式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="referTo"></param>
        /// <returns></returns>
        public static void SetFormula(this Range range, Range referTo) {
            if (referTo == null) {
                return;
            }

            string formula = $"={referTo.Worksheet.Name}!{referTo.Address}";
            range.ClearContents();
            range.Formula = formula;
            range.FormulaHidden = true;
        }
*/

        #endregion GetFormula

        #region Range Address

/*
        /// <summary>
        ///     获取 <see cref="Range">单元格区域</see>地址
        ///     通过 ‘,’ 分隔多个区域地质
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string FullAddress(this Range range) {
            return string.Join(",", range.GetAddress());
        }
*/

/*
        private static IEnumerable<string> GetAddress(this Range range) {
            foreach (Range area in range.Areas) {
                yield return $"{range.Worksheet.Name}!{area.Address}";
            }
        }
*/

        #endregion Range Address
    }
}