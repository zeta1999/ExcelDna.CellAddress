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

namespace ExcelDna{
    /// <summary>
    ///     Convenience struct to work directly with Excel double dates rather than using DateTime.FromOADate() conversions
    ///     for certain calculations there exist a direct and fast operation without going through all the conversions
    ///     Implicit operators make sure that class is equivalent to DateTime and double.
    /// </summary>
    internal struct XlDate : IComparable{
        private double _xlDate;

        #region constructors

        public XlDate(double xlDate){
            _xlDate = xlDate;
        }

        public XlDate(XlDate xlDate){
            _xlDate = xlDate._xlDate;
        }

        public XlDate(DateTime dateTime){
            _xlDate = dateTime.ToOADate();
        }

        public XlDate(int year, int month, int day, int hour = 0, int minute = 0, int second = 0, int millisecond = 0){
            _xlDate = new DateTime(year, month, day, hour, minute, second, millisecond).ToOADate();
        }

        #endregion

        #region properties

        public XlDate Date{
            get { return Math.Floor(_xlDate); }
        }

        public int Year{
            get { return DateTime.FromOADate(_xlDate).Year; }
        }

        public int Month{
            get { return DateTime.FromOADate(_xlDate).Month; }
        }

        public int Day{
            get { return DateTime.FromOADate(_xlDate).Day; }
        }

        public DayOfWeek DayOfWeek{
            get { return DateTime.FromOADate(_xlDate).DayOfWeek; }
        }

        public int DayOfYear{
            get { return DateTime.FromOADate(_xlDate).DayOfYear; }
        }

        public int Hour{
            get { return DateTime.FromOADate(_xlDate).Hour; }
        }

        public int Minute{
            get { return DateTime.FromOADate(_xlDate).Minute; }
        }

        public int Second{
            get { return DateTime.FromOADate(_xlDate).Second; }
        }

        public int Millisecond{
            get { return DateTime.FromOADate(_xlDate).Millisecond; }
        }

        #endregion

        #region Date math

        public XlDate AddMilliseconds(double value){
            return new XlDate(_xlDate + value/86400000.0);
        }

        public XlDate AddSeconds(double value){
            return new XlDate(_xlDate + value/86400.0);
        }

        public XlDate AddMinutes(double value){
            return new XlDate(_xlDate + value/1440.0);
        }

        public XlDate AddHours(double value){
            return new XlDate(_xlDate + value/24.0);
        }

        public XlDate AddDays(double value){
            return new XlDate(_xlDate + value);
        }

        public XlDate AddMonths(int value){
            return new XlDate(DateTime.FromOADate(_xlDate).AddMonths(value));
        }

        public XlDate AddYears(int value){
            return new XlDate(DateTime.FromOADate(_xlDate).AddYears(value));
        }

        #endregion

        #region Operators

        public static double operator -(XlDate lhs, XlDate rhs){
            return lhs._xlDate - rhs._xlDate;
        }

        public static XlDate operator -(XlDate lhs, double rhs){
            lhs._xlDate -= rhs;
            return lhs;
        }

        public static XlDate operator +(XlDate lhs, double rhs){
            lhs._xlDate += rhs;
            return lhs;
        }

        public static XlDate operator +(XlDate d, TimeSpan t){
            var date = new XlDate(d);
            d.AddMilliseconds(t.TotalMilliseconds);
            return date;
        }

        public static XlDate operator -(XlDate d, TimeSpan t){
            var date = new XlDate(d);
            d.AddMilliseconds(-t.TotalMilliseconds);
            return date;
        }

        public static XlDate operator ++(XlDate xDate){
            xDate._xlDate += 1.0;
            return xDate;
        }

        public static XlDate operator --(XlDate xDate){
            xDate._xlDate -= 1.0;
            return xDate;
        }

        public static implicit operator double(XlDate xDate){
            return xDate._xlDate;
        }

        public static implicit operator float(XlDate xDate){
            return (float) xDate._xlDate;
        }

        public static implicit operator XlDate(double xlDate){
            return new XlDate(xlDate);
        }

        public static implicit operator DateTime(XlDate xDate){
            return DateTime.FromOADate(xDate);
        }

        public static implicit operator XlDate(DateTime dt){
            return new XlDate(dt);
        }

        #endregion

        #region formatting

        public override string ToString(){
            return DateTime.FromOADate(_xlDate).ToString();
        }

        public string ToString(string format){
            return DateTime.FromOADate(_xlDate).ToString(format);
        }

        public string ToString(string format, IFormatProvider formatprovider){
            return DateTime.FromOADate(_xlDate).ToString(format, formatprovider);
        }

        #endregion

        #region System

        public int CompareTo(object target){
            if (!(target is XlDate)){
                throw new ArgumentException();
            }

            return (_xlDate).CompareTo(((XlDate) target)._xlDate);
        }

        public override bool Equals(object obj){
            if (obj is XlDate){
                return ((XlDate) obj)._xlDate == _xlDate;
            }
            if (obj is double){
                return ((double) obj) == _xlDate;
            }
            return false;
        }

        public override int GetHashCode(){
            return _xlDate.GetHashCode();
        }

        #endregion
    }
}