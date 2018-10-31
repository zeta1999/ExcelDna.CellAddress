using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Extensions;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna {

    /// <summary>
    /// 单元格地址
    /// </summary>
    public class CellAddress : IEquatable<CellAddress> {

        /// <summary>
        /// Excel 行限制
        /// </summary>
        private const int RowsLimit = 1048576;
        /// <summary>
        /// Excel 列限制
        /// </summary>
        private const int ColumnsLimit = 16384;

        /// <summary>
        /// 单元格数据访问使用 <see cref="ExcelReference"/> 方式
        /// 需要在 ExcelDna 插件环境中使用
        /// </summary>
        public static bool UseExcelReference = true;

        /// <summary>
        /// 不正确的单元格引用
        /// </summary>
        internal const string ErrorReference = "#REF!";

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
            var area = range.Areas.Item[1];
            LocalAddress = area.Address;
            
            Count = area.Cells.Count;
            if (Count == 1) {
                RowFirst = RowLast = area.Row - 1;
                ColumnFirst = ColumnLast = area.Column - 1;
            } else {
                RowFirst = range.Row - 1;
                RowLast = RowFirst + area.Rows.Count - 1;
                ColumnFirst = area.Column - 1;
                ColumnLast = ColumnFirst + area.Columns.Count - 1;
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
            if (rowFirst < 0 && columnFirst < 0) {
                throw new IndexOutOfRangeException("Row or Column out of range");
            }

            if (rowLast < 0 || columnLast < 0) {
                //整行/整列
                if (rowLast < 0) {
                    //整行
                    RowFirst = 0;
                    RowLast = RowsLimit-1;
                    ColumnFirst = columnFirst;
                    ColumnLast = columnLast;
                    Count = Rows * ColumnsLimit;
                } else {
                    //整列
                    RowFirst = rowFirst;
                    RowLast = rowLast;
                    ColumnFirst = 0;
                    ColumnLast = ColumnsLimit-1;
                    Count = Columns * RowsLimit;
                }
                LocalAddress = $"{AddressParser.ToAddress(rowFirst, columnFirst)}:{AddressParser.ToAddress(rowLast, columnLast)}";
            } else {

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
                    Debug.WriteLine($"CellAddress {LocalAddress} SheetId:{_cellReference.SheetId}");
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
                if (string.IsNullOrEmpty(SheetName)) {
                    return LocalAddress.GetHashCode();
                }
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

        

        #endregion

    }
}