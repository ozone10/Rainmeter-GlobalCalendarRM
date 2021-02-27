//-----------------------------------------------------------------------
// <copyright file="Plugin.cs" company="oZone">
// Copyright (c) 2021 oZone
// This file is part of GlobalCalendarRM.
// GlobalCalendarRM is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// You should have received a copy of the GNU General Public License
// along with this program. If not, see https://www.gnu.org/licenses/.
// </copyright>
//-----------------------------------------------------------------------

// System
using System;
using System.Globalization;
using System.Runtime.InteropServices;

// Rainmeter API
using Rainmeter;

namespace Plugin
{
    /// <summary>
    /// Plugin class with <c>Rainmeter</c> API.
    /// </summary>
    public static class Plugin
    {
        /// <summary>
        /// Called when a measure is created (i.e. when a skin is loaded or when
        /// a skin is refreshed). Create your measure object here.
        /// Any other initialization or code that only needs to happen once
        /// should be placed here.
        /// </summary>
        /// <param name="data">
        /// You may allocate and store measure specific data to this variable.
        /// The object you save here will be passed to other functions below.
        /// </param>
        /// <param name="rm">
        /// Internal pointer that is passed to most API functions.
        /// If needed, you may save this value for later use (like for logging
        /// functions).
        /// </param>
        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {
            Measure measure = new Measure();
            data = GCHandle.ToIntPtr(GCHandle.Alloc(measure));
            API api = rm;

            measure.IsLunar = api.ReadInt("Lunar", 0) > 0;
            var calendarName = api.ReadString("Calendar", string.Empty).ToLowerInvariant();
            measure.SelectCalendarType(calendarName);
            measure.DataType = api.ReadString("DataType", string.Empty);
            measure.StrDateTime = api.ReadString("DateTime", string.Empty);
            if (!string.IsNullOrEmpty(measure.StrDateTime))
            {
                try
                {
                    measure.DateTime = DateTime.Parse(measure.StrDateTime);
                }
                catch
                {
                    api.LogF(API.LogType.Warning, "Cannot parse DateTime: '{0}'", measure.StrDateTime);
                    measure.StrDateTime = string.Empty;
                }
            }

            measure.DataTypeValue = measure.SelectDataType();
            if (measure.DataTypeValue < 0)
            {
                api.LogF(API.LogType.Warning, "Selected calendar does not support DataType: '{0}'", measure.DataType);
            }
        }

        /// <summary>
        /// Called by <c>Rainmeter</c> when the measure settings are to be read
        /// directly after <c>Initialize</c>. If "DynamicVariables=1" is set on
        /// the measure, this function is called just before every call to
        /// the Update function during the update cycle.
        /// </summary>
        /// <param name="data">
        /// Pointer to the data set in <c>Initialize</c>.
        /// </param>
        /// <param name="rm">
        /// Internal pointer that is passed to most API functions.
        /// </param>
        /// <param name="maxValue">
        /// Pointer to a double that can be assigned to the default maximum
        /// value for this measure. A value of 0.0 will make it based on
        /// the highest value returned from the Update function.
        /// Do not set maxValue unless necessary.
        /// </param>
        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            _ = rm;
            _ = maxValue;

            Measure measure = data;

            if (string.IsNullOrEmpty(measure.StrDateTime))
            {
                measure.DateTime = DateTime.Now;
                measure.DataTypeValue = measure.SelectDataType();
            }
        }

        /// <summary>
        /// Called by <c>Rainmeter</c> when a measure value is to be updated
        /// (i.e. on each update cycle). The number returned represents
        /// the number value of the measure.
        /// </summary>
        /// <param name="data">
        /// Pointer to the data set in <c>Initialize</c>.
        /// </param>
        /// <returns>
        /// The number value of the measure (as a double).
        /// This value will be used as the string value of the measure
        /// if the GetString function is not used or returns a null.
        /// </returns>
        [DllExport]
        public static double Update(IntPtr data)
        {
            Measure measure = data;
            return measure.DataTypeValue;
        }

        /////// <summary>
        /////// Optional function that returns the string value of the measure.
        /////// Since this function is called 'on-demand' and may be called
        /////// multiple times during the update cycle, do not process any data
        /////// or consume CPU in this function. Do as minimal processing
        /////// as possible to return the desired string. It is recommended to do
        /////// all processing during the Update function and set a string variable
        /////// there and retrieve that string variable in this function.
        /////// The return value must be marshalled from a C# style string
        /////// to a C style string (WCHAR*).
        /////// </summary>
        /////// <param name="data">
        /////// Pointer to the data set in <c>Initialize</c>.
        /////// </param>
        /////// <returns>
        /////// The string value for the measure.
        /////// If you want the number value (returned from Update) to be used
        /////// as the measures value, return null instead.
        /////// The return value must be marshalled.
        /////// </returns>
        ////[DllExport]
        ////public static IntPtr GetString(IntPtr data)
        ////{
        ////    Measure measure = data;
        ////    if (measure.Buffer != IntPtr.Zero)
        ////    {
        ////        Marshal.FreeHGlobal(measure.Buffer);
        ////        measure.Buffer = IntPtr.Zero;
        ////    }

        ////    string stringValue = measure.DataType;
        ////    if (!string.IsNullOrEmpty(stringValue))
        ////    {
        ////        measure.Buffer = Marshal.StringToHGlobalUni(stringValue);
        ////        return measure.Buffer;
        ////    }

        ////    return IntPtr.Zero;
        ////}

        /// <summary>
        /// Called by <c>Rainmeter</c> when a measure is about to be destroyed.
        /// Perform cleanup here.
        /// </summary>
        /// <param name="data">
        /// Pointer to the data set in <c>Initialize</c>.
        /// </param>
        [DllExport]
        public static void Finalize(IntPtr data)
        {
            ////Measure measure = data;
            ////if (measure.Buffer != IntPtr.Zero)
            ////{
            ////    Marshal.FreeHGlobal(measure.Buffer);
            ////}

            GCHandle.FromIntPtr(data).Free();
        }

        /// <summary>
        /// Plugin subclass.
        /// </summary>
        private class Measure
        {
            /////// <summary>
            /////// Gets or sets <c>Buffer</c>.
            /////// Used with <c>GetString</c>.
            /////// </summary>
            ////public IntPtr Buffer { get; set; } = IntPtr.Zero;

            /// <summary>
            /// Gets or sets a value indicating whether lunar calendar should be used.
            /// <c>Rainmeter</c> option Lunar. Use lunar calendar if option value is > 0.
            /// </summary>
            public bool IsLunar { get; set; }

            /// <summary>
            /// Gets or sets <c>DataType</c>.
            /// <c>Rainmeter</c> option DataType. Determine which information should plugin provide.
            /// </summary>
            public string DataType { get; set; }

            /// <summary>
            /// Gets or sets <c>DataTypeValue</c>.
            /// Save <c>SelectDataType()</c> value.
            /// </summary>
            public int DataTypeValue { get; set; }

            /// <summary>
            /// Gets or sets <c>StrDateTime</c>.
            /// Option DateTime. <c>DateTime</c> as string.
            /// </summary>
            public string StrDateTime { get; set; }

            /// <summary>
            /// Gets or sets <c>DateTime</c>.
            /// Option DateTime. Default will use current date.
            /// </summary>
            public DateTime DateTime { get; set; } = DateTime.Now;

            /// <summary>
            /// Gets or sets <c>CalendarType</c>.
            /// <c>Rainmeter</c> option Calendar. Select calendar.
            /// </summary>
            private Calendar CalendarType { get; set; }

            /// <summary>
            /// Gets or sets <c>LunarType</c>.
            /// <c>Rainmeter</c> option Calendar. Only used, when <c>IsLunar</c> is true.
            /// </summary>
            private EastAsianLunisolarCalendar LunarType { get; set; }

            /// <summary>
            /// Creates <c>Measure</c> object. Allow to convert <c>IntPtr</c> to <c>Measure</c>.
            /// </summary>
            /// <param name="data">
            /// Variable to store measure specific data.
            /// </param>
            /// <returns>
            /// Measure object.
            /// </returns>
            public static implicit operator Measure(IntPtr data)
            {
                return (Measure)GCHandle.FromIntPtr(data).Target;
            }

            /// <summary>
            /// Select calendar type.
            /// </summary>
            /// <param name="calendarName">
            /// Option calendar value read from skin.
            /// </param>
            internal void SelectCalendarType(string calendarName)
            {
                if (string.IsNullOrEmpty(calendarName))
                {
                    this.CalendarType = new GregorianCalendar();
                }
                else if (this.IsLunar)
                {
                    EastAsianLunisolarCalendar lunarCalendar;
                    switch (calendarName)
                    {
                        case "japanese":
                            lunarCalendar = new JapaneseLunisolarCalendar();
                            break;

                        case "korean":
                            lunarCalendar = new KoreanLunisolarCalendar();
                            break;

                        case "taiwan":
                            lunarCalendar = new TaiwanLunisolarCalendar();
                            break;

                        default:
                            lunarCalendar = new ChineseLunisolarCalendar();
                            break;
                    }

                    this.LunarType = lunarCalendar;
                    this.CalendarType = this.LunarType;
                }
                else
                {
                    Calendar calendar;
                    switch (calendarName)
                    {
                        case "japanese":
                            calendar = new JapaneseCalendar();
                            break;

                        case "korean":
                            calendar = new KoreanCalendar();
                            break;

                        case "taiwan":
                            calendar = new TaiwanCalendar();
                            break;

                        case "buddhist":
                        case "thai":
                        case "thaibuddhist":
                            calendar = new ThaiBuddhistCalendar();
                            break;

                        case "hebrew":
                            calendar = new HebrewCalendar();
                            break;

                        case "hijiri":
                            calendar = new HijriCalendar();
                            break;

                        case "persian":
                            calendar = new PersianCalendar();
                            break;

                        case "umalqura":
                            calendar = new UmAlQuraCalendar();
                            break;

                        case "julian":
                            calendar = new JulianCalendar();
                            break;

                        default:
                            calendar = new GregorianCalendar();
                            break;
                    }

                    this.CalendarType = calendar;
                }
            }

            /// <summary>
            /// Select DataType to provide information.
            /// </summary>
            /// <returns>
            /// Data as integer.
            /// </returns>
            internal int SelectDataType()
            {
                var dataType = this.DataType.ToLowerInvariant();
                if (string.IsNullOrEmpty(dataType))
                {
                    return -1;
                }

                var dateTime = this.DateTime;
                var calendarType = this.CalendarType;

                if (this.IsLunar)
                {
                    var lunarType = this.LunarType;
                    int sexagenaryYear = lunarType.GetSexagenaryYear(dateTime);

                    switch (dataType)
                    {
                        case "celestialstem":
                            return lunarType.GetCelestialStem(sexagenaryYear);

                        case "sexagenaryyear":
                            return sexagenaryYear;

                        case "terrestrialbranch":
                            return lunarType.GetTerrestrialBranch(sexagenaryYear);
                    }
                }

                int day = calendarType.GetDayOfMonth(dateTime);
                int month = calendarType.GetMonth(dateTime);
                int year = calendarType.GetYear(dateTime);

                switch (dataType)
                {
                    case "month":
                        return month;

                    case "day":
                        return day;

                    case "year":
                        return year;

                    case "era":
                        return calendarType.GetEra(dateTime);

                    case "daysinmonth":
                        return calendarType.GetDaysInMonth(year, month);

                    case "daysinyear":
                        return calendarType.GetDaysInYear(year);

                    case "dayofyear":
                        return calendarType.GetDayOfYear(dateTime);

                    case "dayofweek":
                        return (int)calendarType.GetDayOfWeek(dateTime);

                    case "leapmonth":
                        return calendarType.GetLeapMonth(year);

                    case "monthsinyear":
                        return calendarType.GetMonthsInYear(year);

                    case "weekofyear":
                        return calendarType.GetWeekOfYear(dateTime, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday);

                    case "isleapday":
                        return calendarType.IsLeapDay(year, month, day) ? 1 : 0;

                    case "isleapmonth":
                        return calendarType.IsLeapMonth(year, month) ? 1 : 0;

                    case "isleapyear":
                        return calendarType.IsLeapYear(year) ? 1 : 0;

                    default:
                        return -1;
                }
            }
        }
    }
}
