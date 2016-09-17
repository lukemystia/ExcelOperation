using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ExcelOperation.Common;
using ExcelOperation.Models;
using ExcelOperation.Entity;

namespace ExcelOperation.BizLogic
{
	/// <summary>
	/// 納品書作成
	/// </summary>
	public class MakeDeliverySlip
	{
		/// <summary>
		/// Excelファイルのパス
		/// </summary>
		private static readonly string EXCEL_PATH = C_Path.ORIGINAL_SLIP_DIR_PATH + C_Filename.LEDGER;

		/// <summary>
		/// 対象シート名
		/// </summary>
		private static readonly string SHEET_NAME = C_SheetName.LEDGER;


		/// <summary>
		/// 伝票作成
		/// </summary>
		public void logic()
		{
			try
			{
				CmnLogic.CheckExcelFile(EXCEL_PATH);
				var data = new Delivery(EXCEL_PATH, SHEET_NAME);

				using (var book = new XLWorkbook(EXCEL_PATH, XLEventTracking.Disabled))
				{
					// 雛形シート取得
					var sheet = book.Worksheets
												.Where(x => x.Name == SHEET_NAME)
												.Select(x => x.Worksheet)
												.Single();

					for (data.header.pageNo = 1; data.header.pageNo <= data.header.pageNum; data.header.pageNo++)
					{
						// 雛形シートのコピー 名前:ページ番号
						sheet.CopyTo(data.header.pageNo.ToString());

						// コピーしたシート取得
						var copySheet = book.Worksheets
													.Where(x => x.Name == data.header.pageNo.ToString())
													.Select(x => x.Worksheet)
													.Single();

						// このページで書き込む明細数を決定
						var thisPageWriteNum = CommonExcel.GetWriteNum(data.writeDataList.Count, data.infoData.itemRowNum);

						// 不要範囲消去
						CommonExcel.DelUselessArea(ref copySheet, data.infoData, thisPageWriteNum);

						// データ書き込み
						data.WriteData(ref copySheet, thisPageWriteNum);

						// 情報セルを初期化
						copySheet.Cell(CmnConst.INFO_CELL_ADDRESS).Value = null;

						// 改ページプレビュー設定
						copySheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
					}

					// 雛形シート削除
					book.Worksheets.Delete(SHEET_NAME);

					// 保存するファイル名作成
					var saveFileName = C_Path.MAKE_SLIP_DIR_PATH + SHEET_NAME
													+ "_" + data.header.slipNo
													+ "_" + DateTime.Now.ToLongDateString()
													+ ".xlsx";

					// 別名保存
					book.SaveAs(saveFileName);
				}

				Console.WriteLine("正常終了");

			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
