using System;
using System.Diagnostics;
using System.ComponentModel;
using System.Globalization;
using System.Collections.Generic;

//using DateCalculations;

namespace SBDcalendar
{

    public class SBDCalendarMBOT
    {
        public string monthlongname;
        public string monthshortname;
        public string closedate;
        public int WeekDay;
        public int periodnumber;
        public int periodyear;
        public string closing;
        public string closeday;
        public string daytype;
        public int inputweek;
        public int[] WeeksYear = new int[12];
        public int week;
        public string monthlypayroll;
        public List<int> USpaydays = new List<int>();
        public List<int> PayrollWeeks;

        //Used in GET CLOSING TYPE, CLOSIND DAY & DAY TYPE
        public int ClosingWeeks = 0;
        public List<int> ClosingWeeksList = new List<int>();
        public List<int> PrecloseWeeksList = new List<int>();
        public List<int> MidMonthWeeks1 = new List<int>();
        public List<int> MidMonthWeeks2 = new List<int>();
        public List<int> MidMonthWeeks3 = new List<int>();
        public DateTime dtcurrentdate;
        public int[] WeeksPerMonths = { 5, 4, 4, 5, 4, 4, 5, 4, 4, 5, 4, 4 };
        //Used in GET US PAYROLL DATE
        public string strcurrentYear;
        public int intcurrentYear;

        #region MBOT public classes
        #region GET MONTH LONG NAME
        public string GetPrevMonthLongName(string inputdate, string inputWeeksYear = "")
        {

            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }


            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);
            SBDCalendarMBOT objClosePeriod = new SBDCalendarMBOT();
            objClosePeriod.ClosePeriod(inputdate, daytype, out monthlongname, out monthshortname, out periodnumber, out periodyear);

            return monthlongname;
        }
        #endregion
        #region GET PERIOD YEAR
        public int GetPeriodYear(string inputdate, string inputWeeksYear = "")
        {

            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }


            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);
            SBDCalendarMBOT objClosePeriod = new SBDCalendarMBOT();
            objClosePeriod.ClosePeriod(inputdate, daytype, out monthlongname, out monthshortname, out periodnumber, out periodyear);

            return periodyear;
        }
        #endregion
        #region GET CLOSE DATE
        public string GetCloseDate(string inputdate, string inputWeeksYear = "")
        {

            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }


            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);

            return closedate;
        }
        #endregion
        #region GET MONTH SHORT NAME
        public string GetPrevMonthShortName(string inputdate, string inputWeeksYear = "")
        {
            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }



            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);
            SBDCalendarMBOT objClosePeriod = new SBDCalendarMBOT();
            objClosePeriod.ClosePeriod(inputdate, daytype, out monthlongname, out monthshortname, out periodnumber, out periodyear);

            return monthshortname;
        }
        #endregion
        #region GET PERIOD #
        public int GetPrevMonthPeriodNumber(string inputdate, string inputWeeksYear = "")
        {
            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }
            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);

            SBDCalendarMBOT objClosePeriod = new SBDCalendarMBOT();
            objClosePeriod.ClosePeriod(inputdate, daytype, out monthlongname, out monthshortname, out periodnumber, out periodyear);

            return periodnumber;
        }
        #endregion
        #region CHECK IF ITS CLOSING
        public string CheckClosing(string inputdate, string inputWeeksYear = "")
        {
            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }
            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);

            return closing;
        }
        #endregion
        #region GET CLOSING DAY
        public string GetClosingDay(string inputdate, string inputWeeksYear = "")
        {
            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }
            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);

            return closeday;
        }
        #endregion
        #region GET CLOSING TYPE
        public string GetClosingType(string inputdate, string inputWeeksYear = "")
        {
            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }

            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);

            return daytype;
        }
        #endregion
        #region GET WEEK NUMBER

        public int GetWeekNumber(string inputdate, string inputWeeksYear = "")
        {
            //Parse through inputWeeks string and change it to an int array
            if (inputWeeksYear != "")
            {
                string[] words = inputWeeksYear.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    WeeksYear[i] = inputweek;
                    i++;
                }
            }

            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.CloseType(inputdate, WeeksYear, out closing, out closeday, out daytype, out week, out closedate);

            return week;
        }
        #endregion
        #region CHECK US PAYROLL DATE
        public string CheckUSPayrollDate(string inputdate, string inputUSpaydays = "")
        {
            //Parse through inputWeeks string and change it to an int array
            if (inputUSpaydays != "")
            {
                string[] words = inputUSpaydays.Split(',');
                int i = 0;
                foreach (var word in words)
                {

                    inputweek = Convert.ToInt32(word);
                    USpaydays.Add(inputweek);
                    i++;
                }
            }

            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.GetUSPayrollDate(inputdate, ref USpaydays, out monthlypayroll, out PayrollWeeks);

            return monthlypayroll;
        }
        #endregion
        #region GET PAYROLL WEEKS
        public List<int> GetPayrollWeeks(string inputdate, string inputUSpaydays = "")
        {
            //Parse through inputUSpayweeks string and change it to an int array
            if (inputUSpaydays != "")
            {
                string[] words = inputUSpaydays.Split(',');
                int i = 0;
                foreach (var word in words)
                {
                    inputweek = Convert.ToInt32(word);
                    USpaydays.Add(inputweek);
                    i++;
                }
            }

            SBDCalendarMBOT objCloseType = new SBDCalendarMBOT();
            objCloseType.GetUSPayrollDate(inputdate, ref USpaydays, out monthlypayroll, out PayrollWeeks);

            return PayrollWeeks;
        }
        #endregion
        #endregion
        #region GET MONTH NAME & PERIOD#
        void ClosePeriod(string inputdate, string CloseType, out string LongMonthName, out string shortMonthName, out int periodnumber, out int periodyear)
        {

            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            DateTime dtcurrentdate;

            if (inputdate == "")
            {
                //Setting default parameters for today due to missing input date
                dtcurrentdate = DateTime.Now;
                string strcurrentmonth = DateTime.Now.Month.ToString();
                periodnumber = Convert.ToInt32(strcurrentmonth);
            }
            else
            {
                //Setting  parameters according to input date
                dtcurrentdate = DateTime.Parse(inputdate);
                string strcurrentmonth = dtcurrentdate.Month.ToString();
                periodnumber = Convert.ToInt32(strcurrentmonth);
            }

            //Splitting date variable
            int day = dtcurrentdate.Day;
            int month = dtcurrentdate.Month;
            int year = dtcurrentdate.Year;

            //Get last week of month from current date
            int daysThisMonth = DateTime.DaysInMonth(year, month);
            DateTime endofthismonth = new DateTime(year, month, daysThisMonth);
            int lastweekofmonth = cal.GetWeekOfYear(endofthismonth, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);

            //Get last week number of the year
            DateTime lastDate = new DateTime(year, 12, 31);
            int lastWeek = cal.GetWeekOfYear(lastDate, dfi.CalendarWeekRule, DayOfWeek.Sunday);

            //Get week number from current date
            int currentweek = cal.GetWeekOfYear(dtcurrentdate, dfi.CalendarWeekRule, DayOfWeek.Sunday);

            //Calculate period number
            if (CloseType == "Closing")
            {
                if (currentweek != lastWeek && lastweekofmonth != currentweek && day < 20)
                {
                    periodnumber -= 1;
                }
            }
            else if (CloseType == "PreClose")
            {
                if (day < 15)
                {
                    periodnumber -= 1;
                }
            }
            else if (CloseType == "MidMonth")
            {
                periodnumber -= 1;
            }

            if (periodnumber == 0)
            {
                periodnumber = 12;
            }

            if (periodnumber == 12 && day < 15)
            {
                periodyear = year - 1;
            }
            else
            {
                periodyear = year;
            }

            DateTimeFormatInfo CurrentMonthName = new DateTimeFormatInfo();
            LongMonthName = CurrentMonthName.GetMonthName(periodnumber).ToString(); //Long month name
            shortMonthName = LongMonthName.Substring(0, 3); //Shoth month name

        }
        #endregion
        #region GET CLOSING TYPE, CLOSIND DAY & DAY TYPE
        void CloseType(string inputdate, int[] WeeksYear, out string Closing, out string CloseDay, out string DayType, out int week, out string closedate)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            closedate = "Error";

            if (inputdate == "")
            {
                dtcurrentdate = DateTime.Now;
            }
            else
            {
                dtcurrentdate = DateTime.Parse(inputdate);
            }

            //Check current week number
            week = cal.GetWeekOfYear(dtcurrentdate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday);

            //If custom weeks per month were provided then assign them to the variable
            if (WeeksYear[1] > 0)
            {
                WeeksPerMonths = WeeksYear;
            }

            //Splitting date variable
            int month = dtcurrentdate.Month;
            int year = dtcurrentdate.Year;

            //Get last week of month from current date
            int daysThisMonth = DateTime.DaysInMonth(year, month);
            DateTime lastDate = new DateTime(year, month, daysThisMonth);
            int lastweekofmonth = cal.GetWeekOfYear(lastDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday);

            //Get last week according to WeeksPerMonth
            int sum = 0;
            foreach (int item in WeeksPerMonths)
            {
                sum += item;
            }

            //Setting close types for first weeks of January
            if (week > sum && month == 1)
            {
                PrecloseWeeksList.Add(week);
            }

            ClosingWeeksList.Add(1);
            MidMonthWeeks1.Add(2);
            MidMonthWeeks2.Add(3);
            MidMonthWeeks3.Add(4);

            //Loop through WeeksPerMonths table and check Closing Week in relation to current month number
            for (int index = 0; index < 12; index++)
            {
                ClosingWeeks += WeeksPerMonths[index];
                if (week <= sum || index < 11)
                {
                    ClosingWeeksList.Add(ClosingWeeks + 1);
                    PrecloseWeeksList.Add(ClosingWeeks);
                    MidMonthWeeks1.Add(ClosingWeeks + 2);
                    MidMonthWeeks2.Add(ClosingWeeks + 3);
                    MidMonthWeeks3.Add(ClosingWeeks + 4);
                }
                else if (week > sum && index == 11)
                {

                    if (month == 12)
                    {
                        week = 1;
                        ClosingWeeksList.Add(ClosingWeeks + 1);
                    }
                }
            }

            //Find Prec-close week and Close week on the basis of loop above
            if (PrecloseWeeksList.Contains(week))
            {
                Closing = "True";
                WeekDay = (int)(dtcurrentdate.DayOfWeek - 6);
                if (WeekDay == 0)
                {
                    CloseDay = "Day-" + WeekDay;
                }
                else
                {
                    CloseDay = "Day" + WeekDay;
                }
                DayType = "PreClose";
            }
            else if (ClosingWeeksList.Contains(week))
            {
                Closing = "True";
                WeekDay = (int)(dtcurrentdate.DayOfWeek);
                CloseDay = "Day+" + WeekDay;
                DayType = "Closing";
            }
            else if (MidMonthWeeks1.Contains(week))
            {
                Closing = "False";
                WeekDay = (int)(dtcurrentdate.DayOfWeek + 7);
                CloseDay = "Day+" + WeekDay;
                DayType = "MidMonth";
            }
            else if (MidMonthWeeks2.Contains(week))
            {
                Closing = "False";
                WeekDay = (int)(dtcurrentdate.DayOfWeek + 14);
                CloseDay = "Day+" + WeekDay;
                DayType = "MidMonth";
            }
            else if (MidMonthWeeks3.Contains(week))
            {
                Closing = "False";
                WeekDay = (int)(dtcurrentdate.DayOfWeek + 21);
                CloseDay = "Day+" + WeekDay;
                DayType = "MidMonth";
            }
            else
            {
                Closing = "False";
                CloseDay = "unidentified";
                DayType = "NoType";
            }

            if (CloseDay.Contains("Day-"))
            {
                closedate = dtcurrentdate.AddDays(WeekDay + 1).ToShortDateString();
            }
            else if (CloseDay.Contains("Day+"))
            {
                closedate = dtcurrentdate.AddDays(-(WeekDay + 1)).ToShortDateString();
            }

        }
        #endregion
        #region GET US PAYROLL DATE
        void GetUSPayrollDate(string inputdate, ref List<int> USpaydays, out string MonthlyPayroll, out List<int> PayrollWeeks)
        {
            PayrollWeeks = new List<int>();

            if (inputdate != "")
            {
                //If InputDate has been provided then change it to date
                dtcurrentdate = DateTime.Parse(inputdate);
                strcurrentYear = dtcurrentdate.Year.ToString();
                intcurrentYear = Convert.ToInt32(strcurrentYear);
            }
            else
            {
                //Set default variables
                dtcurrentdate = DateTime.Today;
                strcurrentYear = dtcurrentdate.Year.ToString();
                intcurrentYear = Convert.ToInt32(strcurrentYear);
            }


            //Check current/inputdate week number
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            int week = cal.GetWeekOfYear(dtcurrentdate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday);

            //Find all 2nd Thursdays of the year and add them to a list variable
            if (USpaydays.Count == 0)
            {
                for (int Month = 1; Month <= 12; Month++)
                {
                    DateTime _date = new DateTime(intcurrentYear, Month, 1);
                    DayOfWeek day = _date.DayOfWeek;

                    int d = 0;
                    if (day == DayOfWeek.Saturday || day == DayOfWeek.Friday)
                        d += 7;

                    var diff = DayOfWeek.Thursday - day;

                    DateTime secThursday = _date.AddDays(diff + 7 + d);

                    int PayrollWeek = cal.GetWeekOfYear(secThursday, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);

                    PayrollWeeks.Add(PayrollWeek);
                }
            }
            else
            {
                PayrollWeeks = USpaydays;
            }

            if (PayrollWeeks.Contains(week))
            {
                MonthlyPayroll = "True";
            }
            else
            {
                MonthlyPayroll = "False";
            }


        }
        #endregion

    }

}
