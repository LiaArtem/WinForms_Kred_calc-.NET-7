using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Linq;
using System.Data.OleDb;

namespace Kred_calc
{
	public class Public
	{		
		// Замена сетевых путей
		public static string Korrect_on_web(string in_par)
		{
			string? prom = null;
			in_par = in_par.Trim(' ');
			prom = in_par;
			in_par = in_par.Replace("\\\\", "\\");
			if (prom[..2] == "\\\\")
			{
				in_par = "\\" + in_par;
			}
			return in_par;
		}

        /*
		public static void export_to_excel(string format, dynamic ws_val, int row, int col, string text_t)
		{
			// вывод данных
			if (string.IsNullOrEmpty(text_t))
			{
				return;
			}
			// Числовой
			if (format.ToLower() == "num")
			{
				if (NumericHelper.IsNumeric(text_t) == true)
				{
					ws_val.Cells(row, col) = double.Parse(text_t);
				}
				else
				{
					ws_val.Cells(row, col) = text_t;
				}
				// Текстовый
			}
			else if (format.ToLower() == "text")
			{
				ws_val.Cells(row, col) = text_t;
				// Дата
			}
			else if (format.ToLower() == "date")
			{
				if (DateHelper.IsDate(text_t) == true)
				{
					DateTime zn = Convert.ToDateTime(DateTime.Parse(text_t).ToString("d"));
					if (zn == DateTime.Parse("01.01.0001"))
					{
						ws_val.Cells(row, col) = null;
					}
					else
					{
						ws_val.Cells(row, col) = zn;
					}
				}
				else
				{
					ws_val.Cells(row, col) = text_t;
				}
			}
			else
			{
				if (NumericHelper.IsNumeric(text_t) == true)
				{
					ws_val.Cells(row, col) = double.Parse(text_t);
				}
				else
				{
					ws_val.Cells(row, col) = text_t;
				}
			}
		}
        */
        //Последний день месяца
        public static int LastDayOfMonth(DateTime dteDate)
		{
            return DateTime.DaysInMonth(dteDate.Year, dteDate.Month);			
		}

        // Количество дней в году
		public static int KolDayOfYear(DateTime dteDate)
		{
            DateTime d = new(dteDate.Year, 1, 1);
            DateTime d2 = new(dteDate.Year + 1, 1, 1);
            return Convert.ToInt32((d2 - d).TotalDays);            
		}
	}

}