using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ExcelOperation.Common;
using ExcelOperation.Entity;

namespace ExcelOperation.Models
{
	/// <summary>
	/// 書き込むデータまとめ
	/// </summary>
	public class Delivery
	{
		/// <summary>
		/// 設定情報
		/// </summary>
		public DeliveryInfoData infoData { get; private set; }

		/// <summary>
		/// 明細データ
		/// </summary>
		public List<DeliveryData> writeDataList { get; private set; }

		/// <summary>
		/// ヘッダデータ
		/// </summary>
		public DeliveryHeader header { get; private set; }

		/// <summary>
		/// key:タグ	val:タグのアドレス
		/// </summary>
		private Dictionary<string, string> cellDic;

		/// <summary>
		/// データ取得
		/// </summary>
		/// <param name="filepath">帳票のオリジナルのパス</param>
		/// <param name="sheetName">シート名</param>
		public Delivery(string filepath, string sheetName)
		{
			try
			{
				this.infoData = CommonExcel.GetInfoData<DeliveryInfoData>(filepath, sheetName);

				var itemEntity = new DeliveryItemData();
				this.writeDataList = itemEntity.GetData();

				var headerEntity = new DeliveryHeaderData();
				this.header = headerEntity.GetData(writeDataList, infoData.itemRowNum);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// データ書き込み
		/// </summary>
		/// <param name="sheet">現在のエクセルシート</param>
		/// <param name="writeDataNum">書き込む個数</param>
		public void WriteData(ref IXLWorksheet sheet, int writeDataNum)
		{
			try
			{
				// タグ情報を拾う key=セル値(タグ) val=セルアドレス
				cellDic = CommonExcel.GetTagAddress(sheet);

				WriteItemData(ref sheet, writeDataNum);
				WriteHeader(ref sheet);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 明細タグに書き込む
		/// </summary>
		/// <param name="sheet">現在のエクセルシート</param>
		/// <param name="writeDataNum">書き込む個数</param>
		private void WriteItemData(ref IXLWorksheet sheet, int writeDataNum)
		{
			try
			{
				for (int i = 1; i <= writeDataNum; i++)
				{
					var tag = new DeliveryDataTag(i);

					var temp = writeDataList.First();

					sheet.Cell(cellDic[tag.productName]).Value = temp.productName;
					sheet.Cell(cellDic[tag.size1]).Value = temp.size1;
					sheet.Cell(cellDic[tag.size2]).Value = temp.size2;
					sheet.Cell(cellDic[tag.size3]).Value = temp.size3;
					sheet.Cell(cellDic[tag.qty]).Value = temp.qty;
					sheet.Cell(cellDic[tag.tanka]).Value = temp.tanka;
					sheet.Cell(cellDic[tag.kingaku]).Value = temp.kingaku;
					sheet.Cell(cellDic[tag.bikou]).Value = temp.bikou;

					writeDataList.Remove(temp);
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// ヘッダタグに書き込む
		/// </summary>
		/// <param name="sheet">現在のエクセルシート</param>
		private void WriteHeader(ref IXLWorksheet sheet)
		{
			try
			{
				sheet.Cell(cellDic[DeliveryHedderTag.CustomerName]).Value = header.customerName;
				sheet.Cell(cellDic[DeliveryHedderTag.SlipNo]).Value = header.slipNo;
				sheet.Cell(cellDic[DeliveryHedderTag.SlipDate]).Value = header.slipDate;
				sheet.Cell(cellDic[DeliveryHedderTag.PageNo]).Value = header.pageNo;
				sheet.Cell(cellDic[DeliveryHedderTag.PageNum]).Value = header.pageNum;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}

	/// <summary>
	/// 明細タグ文字列
	/// </summary>
	public class DeliveryDataTag
	{
		/// <summary>
		/// 品名
		/// </summary>
		public string productName { get; private set; }

		/// <summary>
		/// サイズ1
		/// </summary>
		public string size1 { get; private set; }

		/// <summary>
		/// サイズ2
		/// </summary>
		public string size2 { get; private set; }

		/// <summary>
		/// サイズ3
		/// </summary>
		public string size3 { get; private set; }

		/// <summary>
		/// 数量
		/// </summary>
		public string qty { get; private set; }

		/// <summary>
		/// 単価
		/// </summary>
		public string tanka { get; private set; }

		/// <summary>
		/// 金額
		/// </summary>
		public string kingaku { get; private set; }

		/// <summary>
		/// 明細備考
		/// </summary>
		public string bikou { get; private set; }

		/// <summary>
		/// 初期化
		/// </summary>
		/// <param name="i"></param>
		public DeliveryDataTag(int i)
		{
			this.productName = "<品名" + i.ToString() + ">";
			this.size1 = "<サイズ" + i.ToString() + "-1>";
			this.size2 = "<サイズ" + i.ToString() + "-2>";
			this.size3 = "<サイズ" + i.ToString() + "-3>";
			this.qty = "<数量" + i.ToString() + ">";
			this.tanka = "<単価" + i.ToString() + ">";
			this.kingaku = "<金額" + i.ToString() + ">";
			this.bikou = "<備考" + i.ToString() + ">";
		}
	}

	/// <summary>
	/// 明細タグ情報格納クラス
	/// </summary>
	public class DeliveryData
	{
		/// <summary>
		/// 品名
		/// </summary>
		public string productName { get; set; }

		/// <summary>
		/// サイズ1
		/// </summary>
		public int size1 { get; set; }
		
		/// <summary>
		/// サイズ2
		/// </summary>
		public int size2 { get; set; }
		
		/// <summary>
		/// サイズ3
		/// </summary>
		public int size3 { get; set; }

		/// <summary>
		/// 数量
		/// </summary>
		public int qty { get; set; }
		
		/// <summary>
		/// 単価
		/// </summary>
		public int tanka { get; set; }
		
		/// <summary>
		/// 金額
		/// </summary>
		public int kingaku { get; set; }
		
		/// <summary>
		/// 備考
		/// </summary>
		public string bikou { get; set; }
	}

	/// <summary>
	/// ヘッダタグ情報格納クラス
	/// </summary>
	public class DeliveryHeader
	{
		/// <summary>
		/// 伝票No
		/// </summary>
		public string slipNo { get; set; }

		/// <summary>
		/// 伝票日付
		/// </summary>
		public DateTime slipDate { get; set; }
		
		/// <summary>
		/// ページ番号
		/// </summary>
		public int pageNo { get; set; }
		
		/// <summary>
		/// ページ数
		/// </summary>
		public int pageNum { get; set; }
		
		/// <summary>
		/// 取引先名
		/// </summary>
		public string customerName { get; set; }
	}

	/// <summary>
	/// A1セルに書かれている設定格納クラス
	/// </summary>
	public class DeliveryInfoData : InfoData, ISlip
	{
		/// <summary>
		/// 初期化
		/// </summary>
		/// <param name="dic"></param>
		public void Set(Dictionary<string, string> dic)
		{
			try
			{
				this.itemRowStart = dic["明細行"].Split(':')[0];
				this.itemRowEnd = dic["明細行"].Split(':')[1];
				this.itemRowNum = int.Parse(dic["明細行数"]);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
