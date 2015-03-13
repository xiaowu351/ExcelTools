using System;
using System.Collections.Generic;
using System.Text;

namespace CatchExcel
{
	public class ConstDefine
	{
		#region 自定义常量
		public const string ISYES = "是";
		public const string ISNO = "否";
		public const string BHZC1 = "依安全";
		public const string BHZC2 = "依需求";
		public const string BHZC3 = "不需補";
		public const string ZYLY1 = "生產件";
		/// <summary>
		/// 廠內生產件 2014.12.23增加
		/// </summary>
		public const string ZYLY11 = "廠內生產件";
		/// <summary>
		/// 采购件 ：2014.12.23 输出修改为3。原来输出为2
		/// </summary>
		public const string ZYLY2_3 = "採購件";
		public const string ZYLY3 = "托工件";
		public const string ZDGL1 = "逐批領料";
		public const string ZDGL2 = "自動扣料";
		public const string JZFS1 = "現結";
		public const string JZFS2 = "月結";

		/// <summary>
		/// 结束标记
		/// </summary>
		public const string END = "End";

		/// <summary>
		/// 自动编号
		/// </summary>
		public const string ZDBM = "number";

		/// <summary>
		/// 年月日
		/// </summary>
		public const string YYYYMMdd = "yyyyMMdd";

		/// <summary>
		/// 时分秒
		/// </summary>
		public const string HHmmss = "HHmmss";

		/// <summary>
		/// 商品IMSCJ
		/// </summary>
		public const string IMSCJ = "商品IMSCJ";

		/// <summary>
		/// 庫存IMSCL
		/// </summary>
		public const string IMSCL = "庫存IMSCL";

		/// <summary>
		/// 廠商IMSDB
		/// </summary>
		public const string IMSDB = "廠商IMSDB";

		/// <summary>
		/// 客戶IMSEF
		/// </summary>
		public const string IMSEF = "客戶IMSEF";

		/// <summary>
		/// 成本調整IMSCU
		/// </summary>
		public const string IMSCU = "成本調整IMSCU";
		#endregion

		#region Access查詢語句
		/// <summary>
		/// 商品
		/// </summary>
		public static string AccessSqlProduct_imscj = @"select 品號,品名,規格,
			 '' as 固定1,'' as 固定2,'' as 固定3,
			庫存單位,'' as 固定5,'False' as 布尔1,'True' as 布尔2,
			'' as 固定6,'' as 固定7,循環盤點碼,
			'' as 固定8,'' as 固定9,'' as 固定10,'' as 固定11,
			'Z:\111\{0}.jpg' as 图片路径,'' as 固定12,
			主供應商,補貨政策,'0' as 固定13,安全存量,
			'' as 固定14,'' as 固定15,補貨倍量,標準進價,'False' as 布尔3,
			'' as 固定16,'' as 固定17,幣別,'' as 固定18,
			零售價,'False' as 布尔4,
			'' as 固定19,'' as 固定20,'' as 固定21,'' as 固定22,
			'' as 固定23,'' as 固定24,'' as 固定25,'' as 固定26,
			'' as 固定27,'' as 固定28,'' as 固定29,'' as 固定30,
			'' as 固定31,'' as 固定32,'' as 固定33,'' as 固定34,
			儲位,'' as 固定35,'' as 固定36,'' as 固定37,'' as 固定38,
			'' as 固定39,'' as 固定40,'False' as 布尔5,
			'' as 固定41,'' as 固定42,'' as 固定43,'' as 固定44,
			'' as 固定45,'' as 固定46,主要來源,主要庫別,'' as 固定47,'1' as 固定48,
			'' as 固定49,'' as 固定50,'' as 固定51,'' as 固定52,
			'' as 固定53,'' as 固定54,'' as 固定55,'' as 固定56,
			'' as 固定57,'' as 固定58,'' as 固定59,'' as 固定60,
			'' as 固定61,'DS' as 固定62,'' as 固定63,'DS' as 固定64,
			'' as 固定65,'' as 固定66,計劃人員 as 固定67,'' as 固定68,
			'' as 固定69,'' as 固定70,'' as 固定71,'End' from [Sheet1$]  where 品號 {1}''";
		/// <summary>
		/// 库存
		/// </summary>
		public static string AccessSqlProduct_imscl = @"select 'AA' as 固定1,'number{0}',品號,品名,規格,庫存單位,'0' as 固定2,主要庫別,庫存數量, '' as 固定3,'' as 固定4,'' as 固定5,'' as 固定6, '' as 固定7,'' as 固定8,'' as 固定9,'' as 固定10,'' as 固定11,'True' as 布尔1,'' as 固定12,'' as 固定13,'DS' as 固定14,'' as 固定15,'DS' as 固定16,'' as 固定17,'' as 固定18,'End' from [Sheet1$]  where 品號 {1}''";
		/// <summary>
		/// 厂商
		/// </summary>
		public static string AccessSqlProduct_imsdb = @"select 廠商代號,廠商全稱,廠商簡稱,'' as 固定1,'' as 固定2,
負責人,聯系人,電話號碼一,'' as 固定3,傳真,
Email,廠商地址,'' as 固定4,'' as 固定5,'' as 固定6,'' as 固定7,
交易幣別,'2' as 固定8,'1' as 固定9,'' as 固定10,結帳方式,
'' as 固定11,'' as 固定12,'' as 固定13,'' as 固定14,
'' as 固定15,'' as 固定16,'' as 固定17,'DS' as 固定18,'' as 固定19,
'DS' as 固定20,'' as 固定21,'' as 固定22,'' as 固定23,聯系人,'' as 固定24,
廠商地址,'' as 固定25,'' as 固定26,'' as 固定27,'' as 固定28,'' as 固定29,'' as 固定30,
'False' as 布尔1,
'' as 固定31,'' as 固定32,'' as 固定33,'' as 固定34,
'' as 固定35,'' as 固定36,'' as 固定37,'' as 固定38,
'' as 固定39,'' as 固定40,'' as 固定41,'' as 固定42,
'' as 固定43,'' as 固定44,'' as 固定45,'' as 固定46,
'' as 固定47,'' as 固定48,'' as 固定49, '0' as 固定50,
'' as 固定51,'' as 固定52,'' as 固定53,'' as 固定54,'' as 固定55,'' as 固定56,'End' from [Sheet1$]  where 廠商代號 {0}'' ";
		/// <summary>
		/// 客户
		/// </summary>
		public static string AccessSqlProduct_imsef = @"select 客戶代號,客戶全名,客戶簡稱,
'' as 固定1,負責人,聯系人,
'' as 固定2,'' as 固定3,'' as 固定4,'' as 固定5,
'' as 固定6,'' as 固定7,'' as 固定8,'' as 固定9,
'' as 固定10,'' as 固定11, 電話號碼一,'' as 固定88,傳真,Email,
發票地址,發票地址,發票地址,
'' as 固定12,'' as 固定13,'' as 固定14,'' as 固定15,
'' as 固定16,'' as 固定17,'' as 固定18,'2' as 固定19,
'1' as 固定20,交易幣別,
'1' as 固定21,'1' as 固定22,'' as 固定23,'' as 固定24,
結帳方式,'' as 固定26,'' as 固定27,'' as 固定28,
'' as 固定29,'' as 固定30,'' as 固定31,'' as 固定32,
'' as 固定33,'' as 固定34,客戶代號,
'' as 固定35,'' as 固定36,'' as 固定37,'' as 固定38,
'' as 固定39,'' as 固定40,'' as 固定41,'' as 固定42,
'' as 固定43,'' as 固定44,'' as 固定45,'' as 固定46,
'' as 固定47,'' as 固定48,'' as 固定49,'' as 固定50,
'' as 固定51,'' as 固定52,'' as 固定53,'' as 固定54,
'' as 固定55,'' as 固定56,'' as 固定57,'' as 固定58,
'' as 固定59,'' as 固定60,'' as 固定61,'' as 固定62,
'' as 固定63,'' as 固定64,'' as 固定65,'' as 固定66,
'' as 固定67,'' as 固定68,'' as 固定69,'' as 固定70,
'' as 固定71,'' as 固定72,'DS' as 固定73,
'' as 固定74,'' as 固定75,'' as 固定76,'' as 固定77,
'' as 固定78,'' as 固定79,'' as 固定80,'' as 固定81,
'' as 固定82,'' as 固定83,'' as 固定84,'' as 固定85,
'' as 固定86,'' as 固定87,'End' from [Sheet1$]  where 客戶代號 {0}''";

		/// <summary>
		/// 成本调整
		/// </summary>
		public const string AccessSql_imscu = @"select 'AA','number{0}',品號,品名,規格,庫存單位,主要庫別,'' as 固定1,'' as 固定2,成本,'' as 固定3,'' as 固定4,'' as 固定5,'' as 固定6,'' as 固定7,'' as 固定8,'' as 固定9,'True','yyyyMMdd{0}' as 固定10,'yyyyMMdd{0}' as 固定11,'DS' as 固定12,'yyyyMMdd{0}' as 固定13,'DS' as 固定14,'HHmmss{0}' as 固定15,'FixedCost' from [Sheet1$] where 成本 {1} null";
		#endregion
	}
}
