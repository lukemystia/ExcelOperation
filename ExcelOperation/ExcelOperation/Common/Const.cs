using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelOperation.Common
{
	/// <summary>
	/// 固定値
	/// </summary>
	public static class CmnConst
	{
		/// <summary>
		/// 情報セルのアドレス
		/// </summary>
		public static readonly string INFO_CELL_ADDRESS = "AR1";
	}

	/// <summary>
	/// ファイル・フォルダパス
	/// </summary>
	public static class C_Path
	{
		/// <summary>
		/// 帳票オリジナルファイルのフォルダ
		/// </summary>
		public static readonly string ORIGINAL_SLIP_DIR_PATH = @"originalSlipFile\";

		/// <summary>
		/// 作成した帳票ファイルの保存フォルダ
		/// </summary>
		public static readonly string MAKE_SLIP_DIR_PATH = @"makeSlipFile\";
	}

	/// <summary>
	/// ファイル名
	/// </summary>
	public static class C_Filename
	{
		/// <summary>
		/// 帳票
		/// </summary>
		public static readonly string LEDGER = "Ledger.xlsx";
	}

	/// <summary>
	/// シート名
	/// </summary>
	public static class C_SheetName
	{
		/// <summary>
		/// 納品書
		/// </summary>
		public static readonly string LEDGER = "帳票";
	}

	/// <summary>
	/// 納品書ヘッダタグ文字列
	/// </summary>
	public static class DeliveryHedderTag
	{
		/// <summary>
		/// 取引先名
		/// </summary>
		public static readonly string CustomerName = "<取引先名>";

		/// <summary>
		/// 番号
		/// </summary>
		public static readonly string SlipNo = "<番号>";

		/// <summary>
		/// 発行日
		/// </summary>
		public static readonly string SlipDate = "<発行日>";

		/// <summary>
		/// ページ番号
		/// </summary>
		public static readonly string PageNo = "<ページ番号>";

		/// <summary>
		/// ページ数
		/// </summary>
		public static readonly string PageNum = "<ページ数>";
	}
}
