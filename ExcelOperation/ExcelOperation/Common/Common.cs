using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ExcelOperation.Common
{
	/// <summary>
	/// 汎用処理クラス
	/// </summary>
	public static class CmnLogic
	{
		/// <summary>
		/// 文字列から指定正規表現に一致するもののみ取得
		/// </summary>
		/// <param name="cellAddrStr"></param>
		/// <returns></returns>
		public static string GetStr(this string str, string regular)
		{
			Regex re = new Regex(regular);
			return re.Replace(str, "");
		}


		/// <summary>
		/// Excelファイル確認
		/// </summary>
		/// <param name="filePath"></param>
		public static void CheckExcelFile(string filePath)
		{
			try
			{
				if (!File.Exists(filePath)) throw new FileNotFoundException();

				using (var b = new XLWorkbook(filePath, XLEventTracking.Disabled)) { }
			}
			catch (FileNotFoundException ex)
			{
				Console.WriteLine(filePath + "が存在しません");
				throw ex;
			}
			catch (IOException ex)
			{
				Console.WriteLine(filePath + "がすでに開かれています");
				throw ex;
			}
		}

	}
}
