using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelOperation.Models;

namespace ExcelOperation.Entity
{
	/// <summary>
	/// 納品書の明細データ
	/// </summary>
	public class DeliveryItemData
	{
		/// <summary>
		/// データ取得
		/// ダミーデータ
		/// </summary>
		public List<DeliveryData> GetData()
		{
			try
			{
				var writeDataList = new List<DeliveryData>();

				for (int i = 1; i <= 4; i++)
				{
					var data = new DeliveryData();

					var num = i * 100;
					var count = 1;

					data.productName = (num + count++).ToString();

					data.size1 = num + count++;
					data.size2 = num + count++;
					data.size3 = num + count++;

					data.qty = num + count++;
					data.tanka = num + count++;
					data.kingaku = num + count++;

					data.bikou = (num + count++).ToString();

					writeDataList.Add(data);
				}

				return writeDataList;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
