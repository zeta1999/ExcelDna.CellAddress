using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna {
    /// <summary>
    /// 单元格地址解析器
    /// 支持 R1C1 格式和 A1 格式
    /// </summary>
    internal static class AddressParser {
        private static readonly Regex R1C1FormatRegex = new Regex(@".*R(?<R>\d+)C(?<C>\d+)", RegexOptions.Compiled);

        private static readonly Regex A1FormatRegex = new Regex(@"\s*\$?(?<C>[A-Z]+)\$?(?<R>\d+)",
            RegexOptions.Compiled);

        public static bool IsR1C1Format(string address) {
            return R1C1FormatRegex.IsMatch(address);
        }

        public static CellAddress ParseAddress(string address) {
            if (String.IsNullOrEmpty(address)) {
                return null;
            }

            if (address.IndexOf('#') > -1) {
                return CellAddress.Ref;
            }

            if (IsR1C1Format(address)) {
                return ParseAddressR1C1(address);
            }
            return ParseAddressA1(address);
        }

        /// <summary>
        /// 解析 A1 地址格式
        /// "A1" 或者 "$A$1"
        /// </summary>
        /// <param name="a1"></param>
        /// <returns></returns>
        private static CellAddress ParseAddressA1(string a1) {
            int firstRow, firstCol;
            var sheetName = GetSheetName(a1);
            var addressStartIndex = a1.IndexOf('!') + 1;
            var splitIndex = a1.IndexOf(':') + 1;
            if (splitIndex <= 0) {
                //单个单元格
                if (GetRowColForA1(a1, addressStartIndex, a1.Length - addressStartIndex, out firstRow,
                    out firstCol)) {
                    return new CellAddress(sheetName, firstRow - 1, firstCol - 1);
                }
            } else {
                //单元格范围
                if (GetRowColForA1(a1, addressStartIndex, a1.Length - splitIndex, out firstRow,
                        out firstCol) &&
                    GetRowColForA1(a1, splitIndex, a1.Length - splitIndex, out var lastRow,
                        out var lastCol)) {
                    return new CellAddress(sheetName, firstRow - 1, lastRow - 1, firstCol - 1, lastCol - 1);
                }
            }
            return null;
        }

        /// <summary>
        ///     解析 R1C1 地址格式
        /// </summary>
        /// <param name="r1C1"></param>
        /// <returns></returns>
        private static CellAddress ParseAddressR1C1(string r1C1) {
            int firstRow, firstCol;

            var sheetName = GetSheetName(r1C1);
            var addressStartIndex = r1C1.IndexOf('!') + 1;
            var splitIndex = r1C1.IndexOf(':') + 1;
            if (splitIndex <= 0) {
                //单个单元格
                if (GetRowColForR1C1(r1C1, addressStartIndex, r1C1.Length - addressStartIndex, out firstRow,
                    out firstCol)) {
                    return new CellAddress(sheetName, firstRow - 1, firstCol - 1);
                }
            } else {
                if (GetRowColForR1C1(r1C1, addressStartIndex, r1C1.Length - splitIndex, out firstRow,
                        out firstCol) &&
                    GetRowColForR1C1(r1C1, splitIndex, r1C1.Length - splitIndex, out var lastRow,
                        out var lastCol)) {
                    return new CellAddress(sheetName, firstRow - 1, lastRow - 1, firstCol - 1, lastCol - 1);
                }
            }
            return null;
        }

        /// <summary>
        ///     从地址中解析 <see cref="Name" />
        ///     Worksheet 名称以 ‘!’结束
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public static string GetSheetName(string address) {
            if (String.IsNullOrEmpty(address)) {
                throw new ArgumentNullException(nameof(address));
            }
            if (address.LastIndexOf('!') < 0) {
                //地址中不包含 SheetName
                return String.Empty;
            }
            var startIndex = address.IndexOf(']') + 1;
            var endIndex = address.IndexOf('!', startIndex);
            return address.Substring(startIndex, endIndex - startIndex);
        }

        /// <summary>
        ///     解析地址 R1C1格式
        /// </summary>
        /// <param name="address"></param>
        /// <param name="length"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="begining"></param>
        /// <returns></returns>
        private static bool GetRowColForR1C1(string address, int begining, int length, out int row, out int col) {
            var match = R1C1FormatRegex.Match(address, begining, length);
            var result = true;
            result &= Int32.TryParse(match.Groups["R"].Value, out row);
            result &= Int32.TryParse(match.Groups["C"].Value, out col);
            return result;
        }


        /// <summary>
        ///     解析地址 R1C1格式
        /// </summary>
        /// <param name="address"></param>
        /// <param name="length"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="begining"></param>
        /// <returns></returns>
        private static bool GetRowColForA1(string address, int begining, int length, out int row, out int col) {
            var match = A1FormatRegex.Match(address, begining, length);
            var result = true;
            if (match.Success) {
                result &= Int32.TryParse(match.Groups["R"].Value, out row);
                result &= TryParseColumnIndex(match.Groups["C"].Value, out col);
            } else {
                row = -1;
                col = -1;
            }
            return result;
        }

        /// <summary>
        ///     列字符 转换为数值
        /// </summary>
        /// <param name="colStr"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static bool TryParseColumnIndex(string colStr, out int col) {
            if (String.IsNullOrEmpty(colStr)) {
                col = -1;
                return false;
            }
            col = 0;
            foreach (var c in colStr) {
                if (c >= 'A' && c <= 'Z') {
                    col *= 26;
                    col += c - 'A' + 1;
                } else if (c >= 'a' && c <= 'z') {
                    col *= 26;
                    col += c - 'a' + 1;
                }
            }
            return true;
        }

        /// <summary>
        /// 从 0 开始的 行/列索引，计算单元格地址
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        internal static string ToAddress(int rowIndex, int columnIndex) {
            if (rowIndex < 0 || columnIndex < 0) {
                return CellAddress.ErrorReference;
            }
            return $"${GetColumnName(columnIndex)}${rowIndex + 1}";
        }

        internal static string ToAddressR1C1(int rowIndex, int columnIndex) {
            if (rowIndex < 0 || columnIndex < 0) {
                return CellAddress.ErrorReference;
            }
            return $"R{rowIndex+1}C{columnIndex + 1}";
        }
        
        /// <summary>
        /// 获取列名 A~Z AA~ZZ ... ... XFD
        /// A~ZZ 702
        /// </summary>
        /// <param name="colNum">从 0 开始的列索引名称</param>
        /// <returns></returns>
        internal static string GetColumnName(int colNum) {
            if (colNum < 0 || colNum > 16384) {
                return CellAddress.ErrorReference;
            }
            if (colNum < 26) {
                return ((char)('A' + colNum)).ToString();
            }

            const int columnsBound = 26;
            int c1, c2;
            if (colNum < 702) {
                c1 = colNum / columnsBound - 1;
                c2 = colNum % columnsBound;
                return new string(new[] { (char)('A' + c1), (char)('A' + c2) });
            }

            c1 = colNum / (columnsBound * columnsBound) - 1;
            c2 = (colNum % (columnsBound * columnsBound)) / columnsBound - 1;
            var c3 = colNum % columnsBound;

            return new string(new[] { (char)('A' + c1), (char)('A' + c2), (char)('A' + c3) });
        }
    }
}