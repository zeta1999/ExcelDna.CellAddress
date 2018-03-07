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

namespace ExcelDna.Extensions {
    /// <summary>
    ///     Convenience struct to work directly with Excel double dates rather than using DateTime.FromOADate() conversions
    ///     for certain calculations there exist a direct and fast operation without going through all the conversions
    ///     Implicit operators make sure that class is equivalent to DateTime and double.
    /// </summary>
    internal struct XlDate : IComparable {
        private readonly double _xlDate;
        private readonly DateTime _dateTime;

        #region constructors

        public XlDate(double xlDate) {
            _xlDate = xlDate;
            _dateTime = DateTime.FromOADate(xlDate);
        }

        public XlDate(XlDate xlDate) {
            _xlDate = xlDate._xlDate;
            _dateTime = xlDate._dateTime;
        }

        public XlDate(DateTime dateTime) {
            _dateTime = dateTime;
            _xlDate = dateTime.ToOADate();
        }

        public XlDate(int year, int month, int day, int hour = 0, int minute = 0, int second = 0, int millisecond = 0) {
            _dateTime = new DateTime(year, month, day, hour, minute, second, millisecond);
            _xlDate = _dateTime.ToOADate();
        }

        #endregion

        #region properties

        public XlDate Date {
            get { return Math.Floor(_xlDate); }
        }
        /// <summary>获取此实例所表示日期的年份部分。</summary>
        /// <returns>年份（介于 1 和 9999 之间）。</returns>
        public int Year {
            get { return _dateTime.Year; }
        }
        /// <summary>获取此实例所表示日期的月份部分。</summary>
        /// <returns>月组成部分，表示为 1 和 12 之间的一个值。</returns>
        public int Month {
            get { return _dateTime.Month; }
        }
        /// <summary>获取此实例所表示的日期为该月中的第几天。</summary>
        /// <returns>日组成部分，表示为 1 和 31 之间的一个值。</returns>
        public int Day {
            get { return _dateTime.Day; }
        }

        /// <summary>获取此实例所表示的日期是星期几。</summary>
        /// <returns>一个枚举常量，指示此 <see cref="T:System.DateTime" /> 值是星期几。</returns>
        public DayOfWeek DayOfWeek {
            get { return _dateTime.DayOfWeek; }
        }
        /// <summary>获取此实例所表示的日期是该年中的第几天。</summary>
        /// <returns>该年中的第几天，表示为 1 和 366 之间的一个值。</returns>
        public int DayOfYear {
            get { return _dateTime.DayOfYear; }
        }
        /// <summary>获取此实例所表示日期的小时部分。</summary>
        /// <returns>小时组成部分，表示为 0 和 23 之间的一个值。</returns>
        public int Hour {
            get { return _dateTime.Hour; }
        }
        /// <summary>获取此实例所表示日期的分钟部分。</summary>
        /// <returns>分钟组成部分，表示为 0 和 59 之间的一个值。</returns>
        public int Minute {
            get { return _dateTime.Minute; }
        }
        /// <summary>获取此实例所表示日期的秒部分。</summary>
        /// <returns>秒组成部分，表示为 0 和 59 之间的一个值。</returns>
        public int Second {
            get { return _dateTime.Second; }
        }

        /// <summary>获取此实例所表示日期的毫秒部分。</summary>
        /// <returns>毫秒组成部分，表示为 0 和 999 之间的一个值。</returns>
        public int Millisecond {
            get { return _dateTime.Millisecond; }
        }

        #endregion

        #region Date math

        public XlDate AddMilliseconds(double value) {
            return new XlDate(_xlDate + value / 86400000.0);
        }

        public XlDate AddSeconds(double value) {
            return new XlDate(_xlDate + value / 86400.0);
        }

        public XlDate AddMinutes(double value) {
            return new XlDate(_xlDate + value / 1440.0);
        }

        public XlDate AddHours(double value) {
            return new XlDate(_xlDate + value / 24.0);
        }

        public XlDate AddDays(double value) {
            return new XlDate(_xlDate + value);
        }

        public XlDate AddMonths(int value) {
            return new XlDate(_dateTime.AddMonths(value));
        }

        public XlDate AddYears(int value) {
            return new XlDate(_dateTime.AddYears(value));
        }

        #endregion

        #region Operators

        public static double operator -(XlDate lhs, XlDate rhs) {
            return new XlDate(lhs._xlDate - rhs._xlDate);
        }

        public static XlDate operator -(XlDate lhs, double rhs) {
            return new XlDate(lhs._xlDate - rhs);
        }

        public static XlDate operator +(XlDate lhs, double rhs) {
            return new XlDate(lhs._xlDate + rhs);
        }

        public static XlDate operator +(XlDate d, TimeSpan t) {
            return new XlDate(d._dateTime.Add(t));
        }

        public static XlDate operator -(XlDate d, TimeSpan t) {
            return new XlDate(d._dateTime - t);
        }

        public static XlDate operator ++(XlDate xDate) {
            return new XlDate(xDate._xlDate + 1.0);
        }

        public static XlDate operator --(XlDate xDate) {
            return new XlDate(xDate._xlDate - 1.0);
        }

        public static implicit operator double(XlDate xDate) {
            return xDate._xlDate;
        }

        public static implicit operator float(XlDate xDate) {
            return (float)xDate._xlDate;
        }

        public static implicit operator XlDate(double xlDate) {
            return new XlDate(xlDate);
        }

        public static implicit operator DateTime(XlDate xDate) {
            return DateTime.FromOADate(xDate);
        }

        public static implicit operator XlDate(DateTime dt) {
            return new XlDate(dt);
        }

        #endregion

        #region formatting

        public override string ToString() {
            return _dateTime.ToString();
        }

        public string ToString(string format) {
            return _dateTime.ToString(format);
        }

        public string ToString(string format, IFormatProvider formatprovider) {
            return _dateTime.ToString(format, formatprovider);
        }

        #endregion

        #region System

        public int CompareTo(object target) {
            if (!(target is XlDate)) {
                throw new ArgumentException();
            }

            return (_xlDate).CompareTo(((XlDate)target)._xlDate);
        }

        public override bool Equals(object obj) {
            if (obj is XlDate date) {
                return Math.Abs(date._xlDate - _xlDate) <= double.Epsilon;
            }
            if (obj is double) {
                return Math.Abs(((double)obj) - _xlDate) <= double.Epsilon;
            }
            return false;
        }

        public override int GetHashCode() {
            return _xlDate.GetHashCode();
        }

        #endregion
    }
}