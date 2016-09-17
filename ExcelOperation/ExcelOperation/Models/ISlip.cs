using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelOperation.Models
{
	/// <summary>
	/// 伝票
	/// </summary>
	public interface ISlip
	{
		/// <summary>
		/// 設定情報をそれぞれ格納する
		/// </summary>
		/// <param name="infoStr"></param>
		void Set(Dictionary<string, string> infoStr);
	}
}
