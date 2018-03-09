using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using ExcelDna.Extensions;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna {

    /// <summary>
    /// 单元格地址
    /// </summary>
    public class CellAddress : IEquatable<CellAddress> {
        /// <summary>
        /// 不正确的单元格引用
        /// </summary>
        private const string ErrorReference = "#REF!";

        /// <summary>
        /// 错误的地址
        /// </summary>
        public static readonly CellAddress Ref = new CellAddress(ErrorReference);

        private ExcelReference _cellReference;
        private Range _cellRange;

        /// <summary>
        /// CellAddress 析构方法
        /// 销毁内部的 <see cref="Range"/>引用
        /// </summary>
        ~CellAddress() {
            _cellReference = null;
            if (_cellRange != null) {
                try {
                    Marshal.FinalReleaseComObject(_cellRange);
                } catch (InvalidComObjectException ex) {
                    Trace.TraceWarning("Final CellAddress error,{0}",ex.Message);
                } finally {
                    _cellRange = null;
                }
            }
        }


        private CellAddress(string localAddress) {
            LocalAddress = localAddress;
        }

        public CellAddress(ExcelReference reference) : this(reference.SheetNameLocal(),
            reference.RowFirst,reference.RowLast,reference.ColumnFirst,reference.ColumnLast) {
            if (reference.InnerReferences.Count > 1) {
                throw new ArgumentException("CellAddress 只能包括一个区域");
            }
            _cellReference = reference;
        }

        public CellAddress(Range range) {
            _cellRange = range ?? throw new ArgumentNullException(nameof(range));
            if (range.Areas.Count > 1) {
                throw new ArgumentException("Range 只能包括一个区域");
            }
            SheetName = range.Worksheet.Name;
            LocalAddress = range.Address;
            
            Count = range.Count;
            if (Count == 1) {
                RowFirst = RowLast = range.Row - 1;
                ColumnFirst = ColumnLast = range.Column - 1;
            } else {
                RowFirst = range.Row - 1;
                RowLast = RowFirst + range.Rows.Count - 1;
                ColumnFirst = range.Column - 1;
                ColumnLast = ColumnFirst + range.Columns.Count - 1;
            }
        }

        public CellAddress(string sheetName, int rowIndex, int colIndex)
            : this(sheetName, rowIndex, rowIndex, colIndex, colIndex) {
        }

        /// <summary>
        /// 构建一个新的 <see cref="CellAddress"/>
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowFirst">从0开始开始行索引</param>
        /// <param name="rowLast">从0开始最后行索引</param>
        /// <param name="columnFirst">从0开始开始列索引</param>
        /// <param name="columnLast">从0开始的最后列索引</param>
        public CellAddress(string sheetName, int rowFirst, int rowLast, int columnFirst, int columnLast) {
            SheetName = sheetName;
            if (rowFirst < 0 || rowLast < 0 || columnFirst < 0 || columnLast < 0) {
                throw new IndexOutOfRangeException("Row or Column out of range");
            }
            if (columnFirst == columnLast && rowFirst == rowLast) {
                LocalAddress = AddressParser.ToAddress(rowFirst, columnFirst);
            } else {
                LocalAddress =
                    $"{AddressParser.ToAddress(rowFirst, columnFirst)}:{AddressParser.ToAddress(rowLast, columnLast)}";
            }
            RowFirst = rowFirst;
            RowLast = rowLast;
            ColumnFirst = columnFirst;
            ColumnLast = columnLast;
            Count = Rows * Columns;
        }

        #region  public Properties

        /// <summary>
        ///     单元格数量
        /// </summary>
        public int Count { get; }

        /// <summary>
        /// 单元格区域 列数
        /// </summary>
        public int Columns => Math.Abs(ColumnLast - ColumnFirst) + 1;

        /// <summary>
        /// 行数
        /// </summary>
        public int Rows => Math.Abs(RowLast - RowFirst) + 1;

        /// <summary>
        /// 从 0 开始的 首列索引
        /// </summary>
        public int ColumnFirst { get; }

        /// <summary>
        /// 从 0 开始的首行索引
        /// </summary>
        public int RowFirst { get; }

        /// <summary>
        /// 从 0 开始的 最后行索引
        /// </summary>
        public int RowLast { get; }

        /// <summary>
        /// 从 0 开始的最后列索引
        /// </summary>
        public int ColumnLast { get; }

        /// <summary>
        /// 本地单元格地址(不包括 <see cref="SheetName"/>)
        /// </summary>
        public string LocalAddress { get; }

        /// <summary>
        /// 包括 <see cref="SheetName"/>的 单元格地址
        /// </summary>
        public string FullAddress {
            get {
                if (String.IsNullOrEmpty(SheetName)) {
                    return LocalAddress;
                }
                if (SheetName.IndexOf(' ')>-1) {
                    return $"'{SheetName}'!{LocalAddress}";
                }
                return $"{SheetName}!{LocalAddress}";
            }
        }

        /// <summary>
        /// 单元格所在工作表名称
        /// </summary>
        public string SheetName { get; private set; }

        #endregion

        #region Internal Properties

        /// <summary>
        /// 是否包含 <see cref="Range"/>对象实例
        /// </summary>
        internal bool HasRange {
            get { return _cellRange != null; }
        }

        /// <summary>
        /// 是否包含 <see cref="ExcelReference"/> 对象实例
        /// </summary>
        internal bool HasReference {
            get { return _cellReference != null; }
        }



        #endregion

        #region  public methods
        /// <summary>
        /// 单元格COM对象 引用
        /// </summary>
        public Range CellRange {
            get {
                if (_cellRange == null || !Marshal.IsComObject(_cellRange)) {
                    _cellRange = GetRangeImpl();
                }
                return _cellRange;
            }
        }

        /// <summary>
        ///     单元格引用<see cref="ExcelReference" />
        /// </summary>
        public ExcelReference CellReference {
            get {
                if (_cellReference == null) {
                    try {
                        _cellReference = new ExcelReference(RowFirst, RowLast, ColumnFirst, ColumnLast, SheetName);
                    } catch (XlCallException) {
                        _cellReference = new ExcelReference(RowFirst, RowLast, ColumnFirst, ColumnLast);
                        SheetName = _cellReference.SheetNameLocal();
                    }
                    Trace.TraceInformation($"CellAddress {LocalAddress} SheetId:{_cellReference.SheetId}");
                }
                return _cellReference;
            }
        }

        /// <summary>
        /// 解析单元格地址,并返回一个 <see cref="CellAddress"/> 对象
        /// </summary>
        /// <param name="rangeAddress"></param>
        /// <returns></returns>
        public static CellAddress Parse(string rangeAddress) {
            return AddressParser.ParseAddress(rangeAddress);
        }

        #region Methods

        private Range GetRangeImpl() {
            try {
                var xlApp = ExcelDnaUtil.Application;
                if (!(xlApp is Application application)) {
                    throw new NullReferenceException();
                }
                return application.Range[FullAddress];
            } catch (InvalidOperationException ioe) {
                //当前 ExcelApplication 不可用
                Trace.TraceWarning("GetRange Error {0}", ioe);
                throw;
            }
        }

        #endregion


        #region Equality members

        /// <inheritdoc />
        bool IEquatable<CellAddress>.Equals(CellAddress other) {
            return Equals(other);
        }

        /// <summary>
        ///     Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <returns>
        ///     true if the specified object  is equal to the current object; otherwise, false.
        /// </returns>
        /// <param name="obj">The object to compare with the current object. </param>
        public override bool Equals(object obj) {
            if (obj == null) {
                return false;
            }
            if (obj.GetType() != GetType()) {
                return false;
            }
            return GetHashCode() == obj.GetHashCode();
        }

        /// <summary>
        ///     Serves as the default hash function.
        /// </summary>
        /// <returns>
        ///     A hash code for the current object.
        /// </returns>
        public override int GetHashCode() {
            unchecked {
                return SheetName.GetHashCode() ^ LocalAddress.GetHashCode();
            }
        }

        #endregion

        /// <summary>
        /// 返回一个包含<see cref="SheetName"/>和<see cref="LocalAddress"/>的单元格地址
        /// </summary>
        /// <returns></returns>
        public override string ToString() {
            if (String.IsNullOrEmpty(SheetName)) {
                return LocalAddress;
            }
            if (SheetName.Contains(" ")) {
                return $"'{SheetName}'!{LocalAddress}";
            }
            return $"{SheetName}!{LocalAddress}";
        }

        #endregion


        #region implicit

        public static implicit operator CellAddress(string address) {
            if (string.IsNullOrEmpty(address)) {
                return null;
            }
            return AddressParser.ParseAddress(address);
        }

        public static implicit operator CellAddress(ExcelReference reference) {
            if (reference == null) {
                return null;
            }
            return new CellAddress(reference);
        }

        public static implicit operator ExcelReference(CellAddress address) {
            return address?.CellReference;
        }

        /// <summary>
        ///     从Excel 单元格地址字符串 获取 CellAddress 对象实例
        /// 以 “#”开头的地址都为 <see cref="Ref"/>
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public static CellAddress Get(string address) {
            if (address == null) {
                throw new ArgumentNullException(nameof(address));
            }
            if (address.Length == 0) {
                throw new ArgumentException($"Argument Invalid Address {address}");
            }
            return AddressParser.ParseAddress(address);
        }
        #endregion


        #region  Inner Class AddressParser

        /// <summary>
        /// 单元格地址解析器
        /// 支持 R1C1 格式和 A1 格式
        /// </summary>
        private static class AddressParser {
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
                    return Ref;
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
                    int lastRow, lastCol;
                    if (GetRowColForR1C1(r1C1, addressStartIndex, r1C1.Length - splitIndex, out firstRow,
                            out firstCol) &&
                        GetRowColForR1C1(r1C1, splitIndex, r1C1.Length - splitIndex, out lastRow,
                            out lastCol)) {
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
                    return ErrorReference;
                }
                return $"{GetColumnName(columnIndex)}{rowIndex + 1}";
            }

            /// <summary>
            /// 获取列名 A~Z AA~ZZ ... ... XFD
            /// A~ZZ 702
            /// </summary>
            /// <param name="colNum">从 0 开始的列索引名称</param>
            /// <returns></returns>
            private static string GetColumnName(int colNum) {
                if (colNum < 0 || colNum > 16384) {
                    return ErrorReference;
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

        #endregion

    }
}