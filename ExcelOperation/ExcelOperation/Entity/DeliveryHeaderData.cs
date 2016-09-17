using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelOperation.Models;

namespace ExcelOperation.Entity
{
	/// <summary>
	/// 納品書のヘッダデータ
	/// </summary>
	public class DeliveryHeaderData
	{
		/// <summary>
		/// データ取得
		/// ダミーデータ
		/// </summary>
		public DeliveryHeader GetData(List<DeliveryData> writeDataList, int pageMax)
		{
			try
			{
				var data = new DeliveryHeader();

				data.slipNo = "AP-232-0916";
				data.slipDate = DateTime.Now;
				data.customerName = "仮企業";

				data.pageNum = (int)Math.Ceiling((double)writeDataList.Count / (double)pageMax);

				return data;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
