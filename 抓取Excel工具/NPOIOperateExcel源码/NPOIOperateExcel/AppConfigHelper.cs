using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace NPOIOperateExcel
{
    /// <summary>
    /// 读取应用程序的配置文件
    /// </summary>
    public  class AppConfigHelper
    {
        
        /// <summary>
        /// A008客戶下单配件需求读取查询字符串
        /// </summary>
        public static string Str008SqlPart
        {
            get
            {
                //只读字段
                string _str008SqlPart = ConfigurationManager.AppSettings["Access008SqlPart"]; 
                return _str008SqlPart;
            }
        }
        /// <summary>
        /// A008客戶下单訂單數量需求表读取配置  
        /// </summary>
        public static string Str008SqlProduct
        {
            get
            {
                string _str008SqlProduct = ConfigurationManager.AppSettings["Access008SqlProduct"];
                return _str008SqlProduct;
            }
        }
        /// <summary>
        /// A009或者A001BA客戶下单訂單數量需求表读取配置  
        /// </summary>
        public static string Str009SqlProduct
        {
            get
            {
                string _str009SqlProduct = ConfigurationManager.AppSettings["Access009SqlProduct"];
                return _str009SqlProduct;
            }
        }
        /// <summary>
        /// A009或者A001BA客戶下单訂單數量需求表读取配置  
        /// </summary>
        public static string Str009SqlPart
        {
            get
            {
                string _str009SqlPart = ConfigurationManager.AppSettings["Access009SqlPart"];
                return _str009SqlPart;
            }
        }

        /// <summary>
        /// A001KP客戶下单訂單數量需求表读取配置  WG-螺丝包
        /// </summary>
        public static string StrA001KPSqlPart1
        {
            get
            {
                string _strA001KPSqlPart1 = ConfigurationManager.AppSettings["AccessA001KPSqlPart1"];
                return _strA001KPSqlPart1;
            }
        }
        /// <summary>
        /// A001KP客戶下单訂單數量需求表读取配置  KP算料
        /// </summary>
        public static string StrA001KPSqlPart2
        {
            get
            {
                string _strA001KPSqlPart2 = ConfigurationManager.AppSettings["AccessA001KPSqlPart2"];
                return _strA001KPSqlPart2;
            }
        }
        /// <summary>
        /// A001KP客戶下单訂單數量需求表读取配置  大把手算料
        /// </summary>
        public static string StrA001KPSqlPart3
        {
            get
            {
                string _strA001KPSqlPart3 = ConfigurationManager.AppSettings["AccessA001KPSqlPart3"];
                return _strA001KPSqlPart3;
            }
        }
        /// <summary>
        /// A001KP客戶下单訂單數量需求表读取配置  白紅銅-合頁算料
        /// </summary>
        public static string StrA001KPSqlPart4
        {
            get
            {
                string _strA001KPSqlPart4 = ConfigurationManager.AppSettings["AccessA001KPSqlPart4"];
                return _strA001KPSqlPart4;
            }
        }
        /// <summary>
        /// A001KP客戶下单訂單數量需求表读取配置 KP算料  
        /// </summary>
        public static string StrA001KPSqlProduct
        {
            get
            {
                string _strA001KPSqlProduct1 = ConfigurationManager.AppSettings["AccessA001KPSqlProduct"];
                return _strA001KPSqlProduct1;
            }
        }
        /// <summary>
        /// S003BA客戶下单訂單數量需求表读取配置   
        /// </summary>
        public static string StrS003BASqlProduct
        {
            get
            {
                string _strS003BASqlProduct = ConfigurationManager.AppSettings["AccessS003BASqlProduct"];
                return _strS003BASqlProduct;
            }
        }
        /// <summary>
        /// S003BA客戶下单配件數量需求表读取配置  
        /// </summary>
        public static string StrS003BASqlPart
        {
            get
            {
                string _strS003BASqlPart = ConfigurationManager.AppSettings["AccessS003BASqlPart"];
                return _strS003BASqlPart;
            }
        }

        /// <summary>
        /// S003客戶下单訂單數量需求表读取配置   
        /// </summary>
        public static string StrS003SqlProduct
        {
            get
            {
                string _strS003SqlProduct = ConfigurationManager.AppSettings["AccessS003SqlProduct"];
                return _strS003SqlProduct;
            }
        }
        /// <summary>
        /// S003客戶下单配件數量需求表读取配置  
        /// </summary>
        public static string StrS003SqlPart
        {
            get
            {
                string _strS003SqlPart = ConfigurationManager.AppSettings["AccessS003SqlPart"];
                return _strS003SqlPart;
            }
        } 
        /// <summary>
        /// 构造查询零件不含旧编号的条件
        /// </summary>
        /// <returns></returns>
        public static StringBuilder CreateAccessSql()
        {
            StringBuilder sb = new StringBuilder();
            // 零件名称
            sb.AppendFormat(@" select top 1 '螺絲K' as 零件名稱, '規格說明L' as 規格說明,'' as 總用量,'' as 安全庫存量,'' as 舊編號 from [算料$]  UNION ALL ");
            // 输出实际数据
            sb.AppendFormat(" select 螺絲K,  規格說明L,    總用量N,  安全庫存量O,'' from [算料$] where NOT (總用量N = '0')    UNION ALL  ");
            // 零件名称
            sb.AppendFormat(" select top 1 '塑膠',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 膠膜Y,	規格說明Z,		總用量AB,	安全庫存量AC,'' from [算料$]  where  NOT (總用量AB = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '密著卡',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 密著卡AF,	規格說明AG,		總用量AI,	安全庫存量AJ,''  from [算料$]  where  NOT (總用量AI = '0') UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '真空格',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 真空格AM,	規格說明AN,		總用量AP,	安全庫存量AQ,'' from [算料$]  where  NOT (總用量AP = '0') UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '真空格',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 真空格AT,	規格說明AU,		總用量AW,	安全庫存量AX,''  from [算料$]  where  NOT (總用量AW = '0') UNION ALL ");

            //----------------------------------------------------下面的两个条件不同--------
            // 零件名称
            sb.AppendFormat(" select top 1 '氣泡袋',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 氣泡袋BA,	規格說明BB,		總用量BD,	安全庫存量BE,''  from [算料$]  where  總用量BD>0  UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '舒服袋',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 舒服袋BH,	規格說明BI,		總用量BK,	安全庫存量BL,''  from [算料$]  where  總用量BK>0  UNION ALL ");

            //---------------------------------------------------------------------------------
            // 零件名称
            sb.AppendFormat(" select top 1 '塑膠袋',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 塑膠袋BO,	規格說明BP,		總用量BR,	安全庫存量BS ,'' from [算料$]  where  NOT (總用量BR = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '插卡',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 插卡BV,	規格說明BW,		總用量BY,	安全庫存量BZ,''  from [算料$]  where  NOT (總用量BY = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '電腦條碼紙',  '規格說明','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 電腦條碼紙CC,	規格說明CD,		總用量CF,	安全庫存量CG ,'' from [算料$]  where  NOT (總用量CF = '0')  union all  ");
            // 零件名称
            sb.AppendFormat(@" select top 1 '天地板CJ' as 零件名称, '規格說明L' as 规格说明,'' as 总用量,'' as 安全库存量,天地板舊編號CK as 旧编号 from [算料$] UNION ALL ");
            // 输出实际数据
            sb.AppendFormat(" select 天地板CJ,	規格說明CL,	總用量CN	,安全庫存量CO ,	天地板舊編號CK from [算料$] where NOT (總用量CN = '0')    UNION ALL  ");
            // 零件名称
            sb.AppendFormat(" select top 1 '內盒', '規格說明','','', '舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 內盒CR,	規格說明CT,		總用量CV,	安全庫存量CW ,	內盒舊編號CS from [算料$]  where 總用量CV>0 UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '外箱',  '規格說明','','', '舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 外箱CZ,	規格說明DB,		總用量DD,	安全庫存量DE ,	外箱舊編號DA from [算料$]  where 總用量DD>0 ");
            return sb;
        }
        /// <summary>
        /// 构造查询零件不含旧编号的条件
        /// </summary>
        /// <returns></returns>
        public static StringBuilder CreateAccessSqlSave11()
        {
            StringBuilder sb = new StringBuilder();
            // 零件名称
            sb.AppendFormat(@" select top 1 '螺絲K' as 零件名稱, '規格說明L' as 規格說明,'' as 總用量,'' as 安全庫存量 from [算料$]  UNION ALL ");
            // 输出实际数据
            sb.AppendFormat(" select 螺絲K,  規格說明L,  總用量N,  安全庫存量O  from [算料$] where NOT (總用量N = '0')    UNION ALL  ");
            // 零件名称
            sb.AppendFormat(" select top 1 '塑膠',  '規格說明','','' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 膠膜Y,	規格說明Z,	總用量AB,安全庫存量AC  from [算料$]  where  NOT (總用量AB = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '密著卡',  '規格說明','','' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 密著卡AF,規格說明AG,總用量AI,安全庫存量AJ  from [算料$]  where  NOT (總用量AI = '0') UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '真空格', '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 真空格AM,規格說明AN,總用量AP,安全庫存量AQ  from [算料$]  where  NOT (總用量AP = '0') UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '真空格', '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 真空格AT,規格說明AU,總用量AW,安全庫存量AX   from [算料$]  where  NOT (總用量AW = '0') UNION ALL ");

            //----------------------------------------------------下面的两个条件不同--------
            // 零件名称
            sb.AppendFormat(" select top 1 '氣泡袋',  '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 氣泡袋BA,規格說明BB,總用量BD,安全庫存量BE   from [算料$]  where  總用量BD>0  UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '舒服袋',  '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 舒服袋BH,規格說明BI,總用量BK,安全庫存量BL   from [算料$]  where  總用量BK>0  UNION ALL ");

            //---------------------------------------------------------------------------------
            // 零件名称
            sb.AppendFormat(" select top 1 '塑膠袋',  '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 塑膠袋BO,	規格說明BP,	總用量BR,安全庫存量BS  from [算料$]  where  NOT (總用量BR = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '插卡',  '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 插卡BV,規格說明BW,總用量BY,安全庫存量BZ   from [算料$]  where  NOT (總用量BY = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '電腦條碼紙',  '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 電腦條碼紙CC,	規格說明CD,	總用量CF,安全庫存量CG  from [算料$]  where  NOT (總用量CF = '0')  union all  ");
            // 零件名称
            sb.AppendFormat(@" select top 1 '天地板', '規格說明','' ,''   from [算料$] UNION ALL ");
            // 输出实际数据
            sb.AppendFormat(" select 天地板CJ,規格說明CL, 總用量CN,安全庫存量CO  from [算料$] where NOT (總用量CN = '0')    UNION ALL  ");
            // 零件名称
            sb.AppendFormat(" select top 1 '內盒', '規格說明','',''   from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 內盒CR,規格說明CT,	總用量CV,安全庫存量CW   from [算料$]  where 總用量CV>0 UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '外箱',  '規格說明','',''  from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 外箱CZ,規格說明DB,	總用量DD,安全庫存量DE  from [算料$]  where 總用量DD>0 ");
            return sb;
        }
        /// <summary>
        /// 构造查询零件旧编号的条件
        /// </summary>
        /// <returns></returns>
        public static StringBuilder CreateAccessSqlOld()
        {
            StringBuilder sb = new StringBuilder();
            // 零件名称
            sb.AppendFormat(@" select top 1 '天地板CJ' as 零件名称,天地板舊編號CK as 旧编号, '規格說明L' as 规格说明,'' as 用量,'' as 总用量,'' as 安全库存量 from [算料$] UNION ALL ");
            // 输出实际数据
            sb.AppendFormat(" select 天地板CJ,	天地板舊編號CK,	規格說明CL,	用量CM,總用量CN	,安全庫存量CO  from [算料$] where NOT (總用量CN = '0')    UNION ALL  ");
            // 零件名称
            sb.AppendFormat(" select top 1 '內盒', '旧编号', '規格說明','','','' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 內盒CR,	內盒舊編號CS,	規格說明CT,	用量CU,	總用量CV,	安全庫存量CW  from [算料$]  where 總用量CV>0 UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '外箱', '旧编号', '規格說明','','','' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 外箱CZ,	外箱舊編號DA,	規格說明DB,	用量DC,	總用量DD,	安全庫存量DE  from [算料$]  where 總用量DD>0 ");
            return sb;
        }

        /// <summary>
        /// 构造查询零件不含旧编号的条件
        /// </summary>
        /// <returns></returns>
        public static StringBuilder CreateAccessSqlSave()
        {
            StringBuilder sb = new StringBuilder();
            // 零件名称
            sb.AppendFormat(@" select top 1 '螺絲K' as 零件名稱, '規格說明L' as 規格說明,'' as 用量,'' as 總用量,'' as 安全庫存量,'' as 舊編號 from [算料$]  UNION ALL ");
            // 输出实际数据
            sb.AppendFormat(" select 螺絲K,  規格說明L,  用量M,  總用量N,  安全庫存量O,'' from [算料$] where NOT (總用量N = '0')    UNION ALL  ");
            // 零件名称
            sb.AppendFormat(" select top 1 '塑膠',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 膠膜Y,	規格說明Z,	用量AA,	總用量AB,	安全庫存量AC,'' from [算料$]  where  NOT (總用量AB = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '密著卡',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 密著卡AF,	規格說明AG,	用量AH,	總用量AI,	安全庫存量AJ,''  from [算料$]  where  NOT (總用量AI = '0') UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '真空格',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 真空格AM,	規格說明AN,	用量AO,	總用量AP,	安全庫存量AQ,'' from [算料$]  where  NOT (總用量AP = '0') UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '真空格',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 真空格AT,	規格說明AU,	用量AV,	總用量AW,	安全庫存量AX,''  from [算料$]  where  NOT (總用量AW = '0') UNION ALL ");

            //----------------------------------------------------下面的两个条件不同--------
            // 零件名称
            sb.AppendFormat(" select top 1 '氣泡袋',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 氣泡袋BA,	規格說明BB,	用量BC,	總用量BD,	安全庫存量BE,''  from [算料$]  where  總用量BD>0  UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '舒服袋',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 舒服袋BH,	規格說明BI,	用量BJ,	總用量BK,	安全庫存量BL,''  from [算料$]  where  總用量BK>0  UNION ALL ");

            //---------------------------------------------------------------------------------
            // 零件名称
            sb.AppendFormat(" select top 1 '塑膠袋',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 塑膠袋BO,	規格說明BP,	用量BQ,	總用量BR,	安全庫存量BS ,'' from [算料$]  where  NOT (總用量BR = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '插卡',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 插卡BV,	規格說明BW,	用量BX,	總用量BY,	安全庫存量BZ,''  from [算料$]  where  NOT (總用量BY = '0') UNION ALL ");

            // 零件名称
            sb.AppendFormat(" select top 1 '電腦條碼紙',  '規格說明','','','','舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 電腦條碼紙CC,	規格說明CD,	用量CE,	總用量CF,	安全庫存量CG ,'' from [算料$]  where  NOT (總用量CF = '0')  union all  ");
            // 零件名称
            sb.AppendFormat(@" select top 1 '天地板CJ' as 零件名称, '規格說明L' as 规格说明,'' as 用量,'' as 总用量,'' as 安全库存量,天地板舊編號CK as 旧编号 from [算料$] UNION ALL ");
            // 输出实际数据
            sb.AppendFormat(" select 天地板CJ,	規格說明CL,	用量CM,總用量CN	,安全庫存量CO ,	天地板舊編號CK from [算料$] where NOT (總用量CN = '0')    UNION ALL  ");
            // 零件名称
            sb.AppendFormat(" select top 1 '內盒', '規格說明','','','', '舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 內盒CR,	規格說明CT,	用量CU,	總用量CV,	安全庫存量CW ,	內盒舊編號CS from [算料$]  where 總用量CV>0 UNION ALL ");
            // 零件名称
            sb.AppendFormat(" select top 1 '外箱',  '規格說明','','','', '舊編號' from [算料$] union all ");
            // 输出实际数据
            sb.AppendFormat(" select 外箱CZ,	規格說明DB,	用量DC,	總用量DD,	安全庫存量DE ,	外箱舊編號DA from [算料$]  where 總用量DD>0 ");
            return sb;
        }
         
    }
}
