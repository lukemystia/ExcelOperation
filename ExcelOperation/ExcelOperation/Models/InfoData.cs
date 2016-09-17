using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelOperation.Common;

namespace ExcelOperation.Models
{
	/// <summary>
	/// A1セルに書かれている設定共通部
	/// </summary>
	public class InfoData
	{
		/// <summary>
		/// 1つめの明細行開始アドレス
		/// </summary>
		public string itemRowStart { get; protected set; }

		/// <summary>
		/// 1つめの明細行終了アドレス
		/// </summary>
		public string itemRowEnd { get; protected set; }

		/// <summary>
		/// 明細最大行数
		/// </summary>
		public int itemRowNum { get; protected set; }

		/// <summary>
		/// 明細1つの行数を取得
		/// </summary>
		/// <returns></returns>
		public int GetItemRow()
		{
			var tempEnd = int.Parse(this.itemRowEnd.GetStr(@"[^0-9]"));
			var tempStart = int.Parse(this.itemRowStart.GetStr(@"[^0-9]"));

			return tempEnd - tempStart + 1;
		}
	}
}
