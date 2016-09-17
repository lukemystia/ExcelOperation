using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;

using ClosedXML.Excel;
using ExcelOperation.Models;

namespace ExcelOperation.Common
{
	/// <summary>
	/// エクセル操作共通処理
	/// </summary>
	public static  class CommonExcel
	{
		/// <summary>
		/// セル情報取得
		/// </summary>
		/// <typeparam name="T">帳票ごとの設定情報クラス</typeparam>
		/// <param name="filepath">帳票のオリジナルのパス</param>
		/// <param name="sheetName">シート名</param>
		/// <returns></returns>
		public static T GetInfoData<T>(string filepath, string sheetName) where T : class , ISlip, new()
		{
			try
			{
				T infoData = new T();

				// 情報セル取得の為だけにExcelファイルを一回開いて閉じる
				using (var book = new XLWorkbook(filepath, XLEventTracking.Disabled))
				{
					// セルの値取得
					var infoCellData = book.Worksheets
															.Where(x => x.Name == sheetName)
															.Select(x => x.Worksheet.Cell(CmnConst.INFO_CELL_ADDRESS).Value.ToString())
															.Single();

					// key:名前,val:設定値 のDictionary作成
					var dic = infoCellData
										.Substring(1, infoCellData.Length - 2)
										.Split(',')
										.Select(x => x.Split('='))
										.ToDictionary(x => x[0], x => x[1]);

					// 設定情報格納
					infoData.Set(dic);
				}

				return infoData;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}


		/// <summary>
		/// タグアドレス取得
		/// </summary>
		/// <param name="sheet">ワークシート</param>
		/// <returns>key:タグ名 value:アドレスのDictionary</returns>
		public static Dictionary<string, string> GetTagAddress(IXLWorksheet sheet)
		{
			try
			{
				var cellDic = sheet.CellsUsed()
										.Where(x => x.Value.ToString().Substring(0, 1) == "<")
										.Select(x => new { x.Value, x.Address })
										.ToDictionary(x => x.Value.ToString(), x => x.Address.ToString());

				return cellDic;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 現在のページに書き込む明細数取得 と 残り書き込み数の更新
		/// </summary>
		/// <param name="remainderNum">残数</param>
		/// <param name="pageMaxNum">1ページの最大数</param>
		/// <returns></returns>
		public static int GetWriteNum(int remainderNum, int pageMaxNum)
		{
			try
			{
				return (remainderNum - pageMaxNum >= 0) ? pageMaxNum : remainderNum;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 不要な範囲の削除
		/// </summary>
		/// <param name="sheet">現在のエクセルシート</param>
		/// <param name="infoData">設定情報</param>
		/// <param name="writeDataNum">このページで書き込む明細数</param>
		public static void DelUselessArea(ref IXLWorksheet sheet, InfoData infoData, int writeDataNum)
		{
			try
			{
				if (infoData.itemRowNum == writeDataNum) return; // 最大書込みの時は削除不要

				var delStart = infoData.itemRowStart.GetStr(@"[^A-Z]")
								+ (int.Parse(infoData.itemRowStart.GetStr(@"[^0-9]")) + (infoData.GetItemRow() * writeDataNum)).ToString();

				var delEnd = infoData.itemRowEnd.GetStr(@"[^A-Z]")
									+ (int.Parse(infoData.itemRowEnd.GetStr(@"[^0-9]")) + (infoData.GetItemRow() * (infoData.itemRowNum - 1))).ToString();

				sheet.Range(delStart, delEnd).Value = null;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
