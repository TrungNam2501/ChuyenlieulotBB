using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Text;
using System.Windows.Forms;

namespace ChuyenlieulotbbNEW
{
    public partial class Form1 : Form
    {
        private const string TodayFormat = "yyyyMMdd";
        private const string ShortDateFormat = "yyMMdd";
        private const string LotDateFormat = "MMdd";
        private const string OilNoData = "Oil no data available";

        private static readonly string[] MachineIpAddresses =
        {
            "198.1.8.21",
            "198.1.8.22",
            "198.1.8.23",
            "198.1.8.24",
            "198.1.8.35",
            "198.1.8.36",
            "198.1.8.37",
            "198.1.8.38",
            "198.1.8.16",
            "198.1.8.17",
            "198.1.8.15",
            "198.1.8.18"
        };

        public Form1()
        {
            InitializeComponent();
        }
        Dictionary<string, string> cnnstr = new Dictionary<string, string>
        {
            { "V-BB3701", @"Data Source=198.1.8.21;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3702", @"Data Source=198.1.8.22;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3703", @"Data Source=198.1.8.23;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3704", @"Data Source=198.1.8.24;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3705", @"Data Source=198.1.8.35;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3706", @"Data Source=198.1.8.36;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3707", @"Data Source=198.1.8.37;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3708", @"Data Source=198.1.8.38;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "01", @"Data Source=198.1.8.21;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "02", @"Data Source=198.1.8.22;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "03", @"Data Source=198.1.8.23;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "04", @"Data Source=198.1.8.24;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "05", @"Data Source=198.1.8.35;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "06", @"Data Source=198.1.8.36;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "07", @"Data Source=198.1.8.37;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "08", @"Data Source=198.1.8.38;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "33", @"Data Source=198.1.10.33;Initial Catalog=erp;User ID=kendakv2;Password=kenda123" },
            { "33BB", @"Data Source=198.1.10.33;Initial Catalog=BB;User ID=kendakv2;Password=kenda123" },
            { "186", @"Data Source=198.1.9.186;Initial Catalog=InTem;User ID=kendakv2;Password=kenda123" },
            { "maytest", @"Data Source=198.1.10.133;Initial Catalog=LotBB;User ID=kenda;Password=kenda123" },
            { "34", @"Data Source=198.1.10.34;Initial Catalog=P8400;User ID=kendakv2;Password=kenda123" },
            { "V11", @"Data Source=198.1.8.16;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },
            { "V12", @"Data Source=198.1.8.17;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },
            { "V13", @"Data Source=198.1.8.15;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },
            { "V14", @"Data Source=198.1.8.18;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },
            { "4", @"Data Source=198.1.10.4;Initial Catalog=LOT;User ID=kendakv2;Password=kenda123" },




        };
        DataTable dataMail = datatableMail();



        private static DataTable datatableMail()
        {
            DataTable myTable = new DataTable("MachineStatus");

            // Khởi tạo DataColumn cho cột "Máy" và "Trạng thái" với tên cột
            DataColumn machineColumn = new DataColumn("Máy", typeof(string));
            myTable.Columns.Add(machineColumn);

            DataColumn statusColumn = new DataColumn("Trạng thái", typeof(string));
            myTable.Columns.Add(statusColumn);

            DataColumn prodat = new DataColumn("Ngày", typeof(string));
            myTable.Columns.Add(prodat);

            DataColumn soLuong = new DataColumn("Số dòng insert ", typeof(string));
            myTable.Columns.Add(soLuong);

            return myTable;

        }
        static bool PingHost(string ipAddress)
        {
            using (Ping ping = new Ping())
            {
                try
                {
                    PingReply reply = ping.Send(ipAddress);
                    if (reply.Status == IPStatus.Success)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (PingException)
                {
                    return false;
                }
            }
        }

        private string GetConnectionString(string key)
        {
            cnnstr.TryGetValue(key, out string connectionString);
            return connectionString;
        }

        private static string GetMachineCode(string ip)
        {
            if (ip.EndsWith("16"))
                return "V11";
            if (ip.EndsWith("17"))
                return "V12";
            if (ip.EndsWith("15"))
                return "V13";
            if (ip.EndsWith("18"))
                return "V14";

            return "V-BB370" + ip[ip.Length - 1];
        }

        private void AddMailRow(string may, string status, string prodat, int count)
        {
            dataMail.Rows.Add(may, status, prodat, count.ToString());
        }

        private static string BuildSqlIn(IEnumerable<string> values)
        {
            List<string> items = values.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            return items.Count > 0
                ? "('" + string.Join("','", items) + "')"
                : "()";
        }

        private static string GetRubberLevel(string barcode, int? ptype)
        {
            if (ptype.GetValueOrDefault() == 1)
            {
                return "8";
            }

            if (ptype.GetValueOrDefault() == 2)
            {
                return "7";
            }

            if (ptype.HasValue)
            {
                return "";
            }

            if (barcode.Substring(0, 2) == "RR" || barcode.Substring(0, 2) == "RD")
            {
                return "5";
            }

            if (barcode.Substring(0, 2) == "RC")
            {
                return "6";
            }

            return "8";
        }

        private static string BuildMaterialQuery(string planId, string sqlIn, bool allowMaterTypeOil)
        {
            string oilMaterialCondition = allowMaterTypeOil
                ? " AND (Mater_Code NOT LIKE '60%' OR (Mater_Code LIKE '60%' AND (Mater_Code = Mater_Name OR Mater_Type = 1))) "
                : "  AND (     Mater_Code NOT LIKE '60%'     OR (Mater_Code LIKE '60%' AND Mater_Code = Mater_Name)   ) ";

            return
                "WITH ranked_b AS ( " +
                "  SELECT b.*, ROW_NUMBER() OVER ( " +
                "    PARTITION BY b.Barcode, b.mater_code, b.Mater_Type " +
                "    ORDER BY b.SaveTime DESC ) as rn " +
                "  FROM [mfns].[dbo].[Ppt_BarCodeRep] b " +
                "  WHERE b.Plan_ID = '" + planId + "' " +
                "    AND SUBSTRING(b.mater_code, 1, 3) <> '680' " +
                "), " +
                "filtered_b AS ( " +
                "  SELECT * FROM ranked_b WHERE rn = 1 " +
                "), " +
                "c AS ( " +
                "  SELECT " +
                "    b.Mater_Barcode, " +
                "    b.Mater_Code, " +
                "    a.real_weight, " +
                "    CONVERT(varchar, CONVERT(date, b.SaveTime), 120) AS DateColumn, " +
                "    CONVERT(varchar, CONVERT(time, b.SaveTime), 108) AS TimeColumnq, " +
                "    b.Serial_Num, a.Barcode, b.Mater_Name, b.SaveTime, b.Mater_Type " +
                "  FROM [mfns].[dbo].[ppt_weigh] a " +
                "  LEFT JOIN filtered_b b ON a.barcode = b.Barcode " +
                "    AND b.mater_code = a.mater_code " +
                "    AND a.weight_id = b.Mater_Type " +
                "  WHERE a.barcode IS NOT NULL " +
                ") " +
                "SELECT DISTINCT * FROM c " +
                "WHERE Mater_Barcode IS NOT NULL " +
                oilMaterialCondition +
                "  AND Serial_Num IN " + sqlIn + " " +
                "ORDER BY SaveTime;";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region
            //if (PingHost("198.1.8.21"))
            //{
            //    insertBB("V-BB3701");
            //}
            //else
            //{
            //    dataMail.Rows.Add("V-BB3701", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.22"))
            //{
            //    insertBB("V-BB3702");
            //}
            //else
            //{
            //    dataMail.Rows.Add("V-BB3702", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.23"))
            //{
            //    insertBB("V-BB3703");
            //}
            //else
            //{
            //    dataMail.Rows.Add("V-BB3703", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.24"))
            //{
            //    insertBB("V-BB3704");
            //}
            //else
            //{
            //    dataMail.Rows.Add("V-BB3704", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.35"))
            //{
            //    insertBB("V-BB3705");
            //}
            //else
            //{
            //    dataMail.Rows.Add("V-BB3705", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.36"))
            //{
            //    insertBB("V-BB3706");
            //}
            //else
            //{
            //    dataMail.Rows.Add("V-BB3706", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.37"))
            //{
            //    insertBB("V-BB3707");
            //}
            //else
            //{
            //    dataMail.Rows.Add("V-BB3707", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.16"))
            //{
            //    insertHoachat("V11");
            //}
            //else
            //{
            //    dataMail.Rows.Add("HC1", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.17"))
            //{
            //    insertHoachat("V12");
            //}
            //else
            //{
            //    dataMail.Rows.Add("HC2", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.15"))
            //{
            //    insertHoachat("V13");
            //}
            //else
            //{
            //    dataMail.Rows.Add("HC3", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}
            ////
            //if (PingHost("198.1.8.18"))
            //{
            //    insertHoachat("V14");
            //}
            //else
            //{
            //    dataMail.Rows.Add("HC4", "Máy tắt", DateTime.Now.ToString("yyyyMMdd"));
            //}

            #endregion

            //string[] ipAddresses = {  "198.1.8.16", "198.1.8.17", "198.1.8.15", "198.1.8.18" };

            //string[] ipAddresses = { "198.1.8.21", "198.1.8.22", "198.1.8.23", "198.1.8.24", "198.1.8.35", "198.1.8.36", "198.1.8.37", "198.1.8.38" };
            //string[] ipAddresses = { "198.1.8.18" };

            #region coment trước khi bù liệu

            string prodat1 = DateTime.Now.AddDays(-1.0).ToString(TodayFormat);
            string sqlDeletekeochuyentruoc = " delete  [P8400].[dbo].[LOT_PCR] where LEFT(pd_barcode, 2) IN ('RD', 'RR', 'RB', 'RC') and pd_date ='" + prodat1 + "'";
            string cnnInsert34 = GetConnectionString("34");
            bool dellete1 = SqlCnn.ExecuteNonQuery(sqlDeletekeochuyentruoc, cnnInsert34);
            #endregion


            foreach (string ip in MachineIpAddresses)
            {
                string code = GetMachineCode(ip);

                if (PingHost(ip))
                {
                    if (code.StartsWith("V-BB"))
                    {
                        //insertBB(code);
                        insertBBMESMOI(code);
                    }
                    else
                    {
                       insertHoachat(code);
                    }
                }
                else
                {
                    AddMailRow(code, "Máy tắt", DateTime.Now.ToString(TodayFormat), 0);
                }
            }




            mail_Chuyenlieulot(dataMail, DateTime.Now.ToString(TodayFormat));


            Application.Exit();
        }
      

        public void insertHoachat(string may)
        {
            try
            {
                int dem = 0;
                string sql_insertHC = "";
                string prodat1 = DateTime.Now.AddDays(-1).ToString(ShortDateFormat);
                //prodat1 = "260101";
                string prodatsolo = DateTime.Now.AddDays(-1).ToString(LotDateFormat);
                string prodat = DateTime.Now.AddDays(-1).ToString(TodayFormat);
                string planProdat = may + prodat1;
                // planProdat = may + "24043";
                string mesSQL = "select Plan_Id,Recipe_Code,Equip_Code, CONVERT(varchar, CONVERT(date, Start_Date), 120) AS DateColumn,CONVERT(varchar, CONVERT(time, Start_Date), 108) AS TimeColumn,Start_Date" +
                    ",Weight_Man   FROM [CWSS_S7].[dbo].[LR_plan] where Plan_Id like '" + planProdat + "%'";
                string cnnMay = GetConnectionString(may);
                DataTable dataMessql = SqlCnn.ExecuteQuery(mesSQL, cnnMay);
                if (dataMessql.Rows.Count > 0)
                {

                    for (int i = 0; i < dataMessql.Rows.Count; i++)
                    {
                        try
                        {
                            string plan_ID = dataMessql.Rows[i]["Plan_Id"].ToString().Trim();
                            string recipe_Code = dataMessql.Rows[i]["Recipe_Code"].ToString().Trim();
                            string equip_Code = dataMessql.Rows[i]["Equip_Code"].ToString().Trim();
                            string dateColumn = dataMessql.Rows[i]["DateColumn"].ToString().Trim();
                            string timeColumn = dataMessql.Rows[i]["TimeColumn"].ToString().Trim();
                            string startDate = dataMessql.Rows[i]["Start_Date"].ToString().Trim();
                            string weight_Man = dataMessql.Rows[i]["Weight_Man"].ToString().Trim();
                            string solo = "";
                            string ca = GetShift(startDate);
                            string maysolo = "";
                            if (may == "V11")
                            {
                                maysolo = "01";
                            }
                            if (may == "V12")
                            {
                                maysolo = "02";
                            }
                            if (may == "V13")
                            {
                                maysolo = "03";
                            }
                            if (may == "V14")
                            {
                                maysolo = "04";
                            }
                            solo = ca + maysolo + "-" + prodatsolo;

                            string sqlSokgHC = "select a.Plan_id,a.Serial_Num,a.Plan_id + RIGHT('000' + CAST(a.Serial_Num AS VARCHAR(3)), 3) as Barcode," +
                                " a.Real_weight,a.Equip_code,b.Material_Code,b.Real_Weight as weight,b.Weight_Time, CONVERT(varchar, CONVERT(date, b.Weight_Time), 120) AS DateCol" +
                                ",CONVERT(varchar, CONVERT(time, b.Weight_Time), 108) AS TimeCol,Batch_Number as Barscan  " +
                                "from [CWSS_S7].[dbo].[LR_lot] a full join [CWSS_S7].[dbo].[LR_weigh] b on a.Plan_id=b.Plan_id and a.Serial_Num=b.Serial_Num where  a.Plan_id ='" + plan_ID + "'";
                            DataTable dtSqlsokghc = SqlCnn.ExecuteQuery(sqlSokgHC, cnnMay);

                            if (dtSqlsokghc.Rows.Count > 0)
                            {
                                List<DataRow> newRows = new List<DataRow>();
                                foreach (DataRow row in dtSqlsokghc.Rows)
                                {
                                    if (row["Serial_Num"].ToString() == "1")
                                    {
                                        DataRow newRow = dtSqlsokghc.NewRow();

                                        // Copy toàn bộ dữ liệu
                                        newRow.ItemArray = row.ItemArray.Clone() as object[];

                                        // Sửa Barcode
                                        string barcode = row["Barcode"].ToString();

                                        if (barcode.Length >= 3)
                                        {
                                            char[] arr = barcode.ToCharArray();
                                            arr[arr.Length - 3] = '9'; // thay vị trí thứ 3 từ cuối
                                            newRow["Barcode"] = new string(arr);
                                        }

                                        newRows.Add(newRow);
                                    }
                                }

                                // Add vào DataTable
                                foreach (var r in newRows)
                                {
                                    dtSqlsokghc.Rows.Add(r);
                                }

                             

                            }
                            if (dtSqlsokghc.Rows.Count > 0)
                            {
                                for (int j = 0; j < dtSqlsokghc.Rows.Count; j++)
                                {
                                    sql_insertHC += "insert into LOT_PCR values('9','" + dtSqlsokghc.Rows[j]["Barcode"] + "','" + recipe_Code + "','" + dtSqlsokghc.Rows[j]["Real_Weight"] + "'," +
                                        " '" + solo + "','" + dateColumn + "','" + ca + "','" + dtSqlsokghc.Rows[j]["Equip_code"] + "','" + dateColumn + "','" + timeColumn + "'," +
                                        " '" + weight_Man + "','','','','','','','','','','','','','','','','" + dtSqlsokghc.Rows[j]["Barscan"] + "'," +
                                        " '" + dtSqlsokghc.Rows[j]["Material_Code"] + "','" + dtSqlsokghc.Rows[j]["weight"] + "','" + dtSqlsokghc.Rows[j]["DateCol"] + "'," +
                                        " '" + dtSqlsokghc.Rows[j]["TimeCol"] + "','" + weight_Man + "');";
                                    dem = dem + 1;
                                }
                            }
                        }
                        catch { continue; }

                    }
                    string cnnInsert34 = GetConnectionString("34");
                    bool a = SqlCnn.ExecuteNonQuery(sql_insertHC, cnnInsert34);
                    string cnnInsert4 = GetConnectionString("4");
                    bool b = SqlCnn.ExecuteNonQuery(sql_insertHC, cnnInsert4);


                    if (a && b)
                    {
                        AddMailRow(may, "Thành công máy hóa chất ", prodat, dem);
                    }
                    else
                    {
                        AddMailRow(may, "Chưa insert được máy hóa chất", prodat, dem);

                    }
                }
                else
                {
                    AddMailRow(may, "Máy hôm qua không chạy", prodat, dem);
                }

            }
            catch (Exception)
            {
                AddMailRow(may, "Văng Exception ex hóa chất , máy có thể tắt ", DateTime.Now.AddDays(-1).ToString(TodayFormat), 0);
            }


        }
        static string GetShift(string timeString)
        {
            // Ép kiểu chuỗi sang kiểu DateTime
            DateTime time = DateTime.ParseExact(timeString, "yyyy-MM-dd HH:mm:ss", null);

            // Xác định giờ và phút
            int hour = time.Hour;
            int minute = time.Minute;

            // Kiểm tra xem thời gian nằm trong khoảng nào
            if ((hour > 6 || (hour == 6 && minute >= 30)) && (hour < 18 || (hour == 18 && minute < 30)))
            {
                return "1"; // Khoảng thời gian từ 6h30 sáng đến 18h30 chiều
            }
            else
            {
                return "2"; // Khoảng thời gian từ 18h30 chiều đến 6h30 sáng hôm sau
            }
        }
        public string Laytemquet(string materialCode, string weightTime, string cnnMay)
        {
            try
            {
                string barcodeQuetsudung = "";
                string sqlBarcodequet = "SELECT top 1 Material,Scan_Bar,Scan_Time FROM [CWSS_S7].[dbo].[LR_BarcodeLog] where  " +
                    "SUBSTRING(convert(varchar,Scan_bar),1,5) ='" + materialCode + "'  and CONVERT(datetime, [Scan_Time], 102) <  CONVERT(DATETIME, '" + weightTime + "', 102)" +
                    " order by Scan_Time desc";
                DataTable dtsqlBarcodequet = SqlCnn.ExecuteQuery(sqlBarcodequet, cnnMay);
                if (dtsqlBarcodequet.Rows.Count > 0)
                {
                    barcodeQuetsudung = dtsqlBarcodequet.Rows[0]["Scan_Bar"].ToString().Trim();
                    return barcodeQuetsudung;
                }
                else
                {
                    return barcodeQuetsudung;
                }


            }
            catch (Exception)
            {
                return "";
            }
        }
     

        public void insertBB(string may)
        {
            try
            {
                int dem = 0;
                string prodat1 = DateTime.Now.AddDays(-1.0).ToString(TodayFormat);
                string SQLVlevel = "SELECT ptype,rubno_7 FROM [InTem].[dbo].[rubnod_Ptype]";
                string cnn186 = GetConnectionString("186");
                DataTable DTcheckvlevel = SqlCnn.ExecuteQuery(SQLVlevel, cnn186);
                string sql_insert = "";
                string cnn33 = GetConnectionString("33");
                string sql = "SELECT mesid,barcode,partno,weight,slipno,prodat,class,machno,indat,intime,usrno,some_sx  FROM [erp].[dbo].[prdebe] where machno ='" + may + "' and prodat  like '" + prodat1 + "%' order by indat , intime ";
                //string sql = "SELECT mesid,barcode,partno,weight,slipno,prodat,class,machno,indat,intime,usrno,some_sx  FROM [erp].[dbo].[prdebe] where machno ='" + may + "' and partno like '755%' and mesid !='RL' and len(mesid) <14 and " +
                    //"barcode IN ('RR26309043','RR26228014','RC26123043','RD26120070','RR25C18033','RR25B19003','RC25423012','RB25422076','RR24613008','RR24611012','RR24611013','RC23713266','RR23708014','RC23629068','RC23629099','RC23629102')  order by indat , intime ";
                DataTable dt = SqlCnn.ExecuteQuery(sql, cnn33);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            string mes = dt.Rows[i]["mesid"].ToString().Trim();
                            string barcode = dt.Rows[i]["barcode"].ToString().Trim();
                            string pArtno = dt.Rows[i]["partno"].ToString().Trim();
                            string wEight = dt.Rows[i]["weight"].ToString().Trim();
                            string sLipno = dt.Rows[i]["slipno"].ToString().Trim();
                            string pRodat = dt.Rows[i]["prodat"].ToString().Trim();
                            string cLass = dt.Rows[i]["class"].ToString().Trim();
                            string mAchno = dt.Rows[i]["machno"].ToString().Trim();
                            string iNdat = dt.Rows[i]["indat"].ToString().Trim();
                            string iNtime = dt.Rows[i]["intime"].ToString().Trim();
                            string uSrno = dt.Rows[i]["usrno"].ToString().Trim();
                            string sMesx = dt.Rows[i]["some_sx"].ToString().Trim();
                            int? ptype = GetPtypeFromRubno7(DTcheckvlevel, pArtno);
                            string level = GetRubberLevel(barcode, ptype);
                            DataTable dtDatangguyenlieu = new DataTable();
                            string sql186 = "SELECT  [machno],[idGrouplot] FROM [InTem].[dbo].[KEORE] where mesid ='" + mes + "'";
                            DataTable dtIDGrouplot = SqlCnn.ExecuteQuery(sql186, cnn186);
                            if (dtIDGrouplot.Rows.Count > 0)
                            {
                                string idGrouplot = dtIDGrouplot.Rows[0][1].ToString().Trim();
                                string sqlPlan_id = "SELECT  [Plan_ID] FROM [mfns].[dbo].[Ppt_GroupLot]  where id='" + idGrouplot + "'";
                                string cnnMay = GetConnectionString(may.Trim().Substring(may.Length - 2));
                                DataTable dtPlanid = SqlCnn.ExecuteQuery(sqlPlan_id, cnnMay);
                                if (dtPlanid.Rows.Count > 0)
                                {
                                    string meso1 = "";
                                    string meso2 = "";
                                    string meso1sau = "";
                                    string meso2sau = "";
                                    string plan_id = dtPlanid.Rows[0][0].ToString().Trim();
                                    //cnnstr.TryGetValue("33", out var cnnSome);
                                    //string sqlSome = "SELECT  barcode,(ROW_NUMBER() OVER (ORDER BY indat, intime) * 2 - 1) AS row_number_1, (ROW_NUMBER() OVER (ORDER BY indat, intime) * 2) AS row_number_2 FROM [erp].[dbo].[prdebe] WHERE mesid = '" + mes + "' ORDER BY indat, intime;";
                                    //DataTable dtSome = SqlCnn.ExecuteQuery(sqlSome, cnnSome);
                                    //if (dtSome.Rows.Count > 0)
                                    //{
                                    //    foreach (DataRow item in dtSome.Rows)
                                    //    {
                                    //        if (item["barcode"].ToString().Trim() == barcode.Trim())
                                    //        {
                                    //            meso1 = item["row_number_1"].ToString().Trim();
                                    //            meso2 = item["row_number_2"].ToString().Trim();
                                    //            break;
                                    //        }
                                    //    }
                                    //}

                                    if (sMesx.Contains("-"))
                                    {
                                        string[] parts = sMesx.Split('-');
                                        meso1 = parts[0];
                                        meso2 = parts[0];
                                        meso1sau = parts[0]; ;
                                        meso2sau = parts[0]; ;

                                    }
                                    else
                                    {
                                        meso1 = sMesx;
                                        meso2 = "";
                                        meso1sau = sMesx;
                                        meso2sau = "";
                                    }
                                    //string dataNguyenlieu = "WITH c AS (SELECT b.Mater_Barcode, b.Mater_Code,a.real_weight, CONVERT(varchar, CONVERT(date, b.SaveTime), 120) AS DateColumn,CONVERT(varchar, CONVERT(time, b.SaveTime), 108) AS TimeColumnq,b.Serial_Num,a.Barcode  FROM [mfns].[dbo].[ppt_weigh] a FULL JOIN [mfns].[dbo].[Ppt_BarCodeRep] b ON a.barcode =b.Barcode AND b.mater_code = a.mater_code WHERE b.Plan_ID = '" + plan_id + "'AND  substring(b.mater_code, 1, 3) <> '680' AND a.barcode IS NOT NULL ) SELECT DISTINCT *FROM c WHERE  Serial_Num in ('" + meso1 + "','" + meso2 + "') order by Serial_Num";
                                    //dtDatangguyenlieu = SqlCnn.ExecuteQuery(dataNguyenlieu, cnnMay);
                                    //if (dtDatangguyenlieu.Rows.Count == 0)
                                    //{
                                    //    dataNguyenlieu = "WITH c AS (SELECT b.Mater_Barcode, b.Mater_Code,a.real_weight, CONVERT(varchar, CONVERT(date, b.SaveTime), 120) AS DateColumn,CONVERT(varchar, CONVERT(time, b.SaveTime), 108) AS TimeColumnq,b.Serial_Num,a.Barcode  FROM [mfns].[dbo].[ppt_weigh] a FULL JOIN [mfns].[dbo].[Ppt_BarCodeRep] b ON a.barcode =b.Barcode AND b.mater_code = a.mater_code WHERE b.Plan_ID = '" + plan_id + "'AND  substring(b.mater_code, 1, 3) <> '680' AND a.barcode IS NOT NULL ) SELECT DISTINCT *FROM c WHERE  Serial_Num in ('" + (int.Parse(meso1) - 1) + "','" + (int.Parse(meso2) - 1) + "')";
                                    //    dtDatangguyenlieu = SqlCnn.ExecuteQuery(dataNguyenlieu, cnnMay);
                                    //    if (dtDatangguyenlieu.Rows.Count == 0)
                                    //    {
                                    //        dataNguyenlieu = "WITH c AS (SELECT b.Mater_Barcode, b.Mater_Code,a.real_weight, CONVERT(varchar, CONVERT(date, b.SaveTime), 120) AS DateColumn,CONVERT(varchar, CONVERT(time, b.SaveTime), 108) AS TimeColumnq,b.Serial_Num,a.Barcode  FROM [mfns].[dbo].[ppt_weigh] a FULL JOIN [mfns].[dbo].[Ppt_BarCodeRep] b ON a.barcode =b.Barcode AND b.mater_code = a.mater_code WHERE b.Plan_ID = '" + plan_id + "'AND  substring(b.mater_code, 1, 3) <> '680' AND a.barcode IS NOT NULL ) SELECT DISTINCT *FROM c WHERE   Serial_Num in ('" + (int.Parse(meso1) - 2) + "','" + (int.Parse(meso2) - 2) + "')";
                                    //        dtDatangguyenlieu = SqlCnn.ExecuteQuery(dataNguyenlieu, cnnMay);
                                    //    }
                                    //}

                                    for (int c = 0; c <= 2; c++)
                                    {
                                        //string query = "WITH c AS " +
                                        //    "(SELECT b.Mater_Barcode, b.Mater_Code,a.real_weight, " +
                                        //    "CONVERT(varchar, CONVERT(date, b.SaveTime), 120) AS DateColumn, " +
                                        //    "CONVERT(varchar, CONVERT(time, b.SaveTime), 108) AS TimeColumnq, " +
                                        //    "b.Serial_Num,a.Barcode " +
                                        //    "FROM [mfns].[dbo].[ppt_weigh] a FULL JOIN [mfns].[dbo].[Ppt_BarCodeRep] b " +
                                        //    "ON a.barcode = b.Barcode AND b.mater_code = a.mater_code " +
                                        //    "WHERE b.Plan_ID = '" + plan_id + "' AND substring(b.mater_code, 1, 3) <> '680' " +
                                        //    "AND a.barcode IS NOT NULL ) " +
                                        //    "SELECT DISTINCT * FROM c WHERE Serial_Num in ('" + meso1 + "','" + meso2 + "') " +
                                        //    "ORDER BY Serial_Num";

                                        string query = BuildMaterialQuery(plan_id, "('" + meso1 + "','" + meso2 + "')", false);

                                        dtDatangguyenlieu = SqlCnn.ExecuteQuery(query, cnnMay);

                                        if (dtDatangguyenlieu.Rows.Count > 0)
                                            break;

                                        // giảm đi 1 cho lần tiếp theo
                                        meso1 = (int.Parse(meso1) - 1).ToString();

                                        if (!string.IsNullOrEmpty(meso2))
                                            meso2 = (int.Parse(meso2) - 1).ToString();
                                    }

                                    if (dtDatangguyenlieu.Rows.Count == 0)
                                    {
                                        meso1 = meso1sau;
                                        meso2 = meso2sau;

                                        for (int c = 0; c <= 2; c++)
                                        {
                                            //string query = "WITH c AS " +
                                            //    "(SELECT b.Mater_Barcode, b.Mater_Code,a.real_weight, " +
                                            //    "CONVERT(varchar, CONVERT(date, b.SaveTime), 120) AS DateColumn, " +
                                            //    "CONVERT(varchar, CONVERT(time, b.SaveTime), 108) AS TimeColumnq, " +
                                            //    "b.Serial_Num,a.Barcode " +
                                            //    "FROM [mfns].[dbo].[ppt_weigh] a FULL JOIN [mfns].[dbo].[Ppt_BarCodeRep] b " +
                                            //    "ON a.barcode = b.Barcode AND b.mater_code = a.mater_code " +
                                            //    "WHERE b.Plan_ID = '" + plan_id + "' AND substring(b.mater_code, 1, 3) <> '680' " +
                                            //    "AND a.barcode IS NOT NULL ) " +
                                            //    "SELECT DISTINCT * FROM c WHERE Serial_Num in ('" + meso1 + "','" + meso2 + "') " +
                                            //    "ORDER BY Serial_Num";

                                            string query = BuildMaterialQuery(plan_id, "('" + meso1 + "','" + meso2 + "')", false);

                                            dtDatangguyenlieu = SqlCnn.ExecuteQuery(query, cnnMay);

                                            if (dtDatangguyenlieu.Rows.Count > 0)
                                                break;

                                            // giảm đi 1 cho lần tiếp theo
                                            meso1 = (int.Parse(meso1) + 1).ToString();

                                            if (!string.IsNullOrEmpty(meso2))
                                                meso2 = (int.Parse(meso2) + 1).ToString();
                                        }

                                    }


                                    string m1 = meso1;
                                    string m2 = string.IsNullOrEmpty(meso2) ? "0" : meso2;
                                    string dataDau = "  SELECT a.[barcode],a.[mater_code], a.[equip_code], a.[set_weight], CONVERT(nvarchar(20),a.[weigh_time],120) as weigh_time ,a.[error_allow], a.[weigh_type], b.mater_name  FROM[mfns].[dbo].[ppt_weigh] a, [mfns].[dbo].[pmt_material] b where barcode in('" + plan_id + int.Parse(m1).ToString("000") + "','" + plan_id + int.Parse(m2).ToString("000") + "') and weigh_type='油料' and a.mater_code = b.mater_code  order by weigh_time asc ";
                                    DataTable dtDatangguyenlieuDau = SqlCnn.ExecuteQuery(dataDau, cnnMay);
                                    if (dtDatangguyenlieuDau.Rows.Count != 0)
                                    {
                                        string coal_barcode = "";
                                        string s_fromday = dtDatangguyenlieuDau.Rows[0]["weigh_time"].ToString().Trim();
                                        DateTime dat_point = DateTime.Parse(s_fromday);
                                        DateTime dat_check = DateTime.Parse(s_fromday);
                                        TimeSpan ts_check = new TimeSpan(6, 30, 0);
                                        dat_check = dat_check.Date + ts_check;
                                        if (dat_point <= dat_check)
                                        {
                                            dat_check = dat_check.AddDays(-1.0);
                                        }
                                        string coal_code = "";
                                        for (int _icount = 0; _icount < dtDatangguyenlieuDau.Rows.Count; _icount++)
                                        {
                                            if (dtDatangguyenlieuDau.Rows[_icount]["weigh_type"].ToString().Trim() == "油料")
                                            {
                                                coal_code = dtDatangguyenlieuDau.Rows[_icount]["mater_code"].ToString().Trim();
                                                break;
                                            }
                                        }
                                        string s_coal_barcode = "  SELECT top 1 [Mater_Barcode],[SaveTime] FROM [mfns].[dbo].[Ppt_Oil] where  SaveTime <= '" + s_fromday + "' and Mater_Type = '0' and [Mater_Code] ='" + coal_code + "' order by SaveTime desc";
                                        DataTable dt_coal_barcode = SqlCnn.ExecuteQuery(s_coal_barcode, cnnMay);
                                        if (dt_coal_barcode.Rows.Count != 0)
                                        {
                                            coal_barcode = dt_coal_barcode.Rows[0][0].ToString().Trim();
                                        }
                                        int i_max = dtDatangguyenlieuDau.Rows.Count;
                                        for (int i_count = 0; i_count < i_max; i_count++)
                                        {
                                            DataRow dr = dtDatangguyenlieu.NewRow();
                                            if (coal_barcode.ToString() != "")
                                            {
                                                if (dtDatangguyenlieuDau.Rows[i_count]["mater_code"].ToString().Trim() == coal_barcode.Substring(0, 5))
                                                {
                                                    dr[0] = coal_barcode;
                                                }
                                                else
                                                {
                                                    dr[0] = OilNoData;
                                                }
                                            }
                                            else
                                            {
                                                dr[0] = OilNoData;
                                            }
                                            dr[1] = dtDatangguyenlieuDau.Rows[0][1].ToString().Trim();
                                            dr[2] = dtDatangguyenlieuDau.Rows[i_count]["set_weight"].ToString();
                                            dr[3] = dtDatangguyenlieuDau.Rows[i_count]["weigh_time"].ToString().Trim().Substring(0, 10);
                                            dr[4] = dtDatangguyenlieuDau.Rows[i_count]["weigh_time"].ToString().Trim().Substring(11);
                                            dr[5] = int.Parse(dtDatangguyenlieuDau.Rows[i_count]["barcode"].ToString().Trim().Substring(dtDatangguyenlieuDau.Rows[i_count]["barcode"].ToString().Trim().Length - 2, 2));
                                            dr[6] = dtDatangguyenlieuDau.Rows[i_count]["barcode"].ToString();
                                            dtDatangguyenlieu.Rows.Add(dr);
                                        }
                                    }
                                }
                            }
                            if (dtDatangguyenlieu.Rows.Count > 0)
                            {
                                for (int j = 0; j < dtDatangguyenlieu.Rows.Count; j++)
                                {
                                    sql_insert = sql_insert + "insert into LOT_PCR values('" + level + "','" + barcode + "','" + pArtno + "','" + wEight + "', '" + sLipno + "','" + pRodat + "','" + cLass + "','" + mAchno + "','" + iNdat + "','" + iNtime + "', '" + uSrno + "','','','','','','','','','','','','','','','','" + dtDatangguyenlieu.Rows[j]["Mater_Barcode"]?.ToString() + "', '" + dtDatangguyenlieu.Rows[j]["Mater_Code"]?.ToString() + "','" + dtDatangguyenlieu.Rows[j]["real_weight"]?.ToString() + "','" + dtDatangguyenlieu.Rows[j]["DateColumn"]?.ToString() + "', '" + dtDatangguyenlieu.Rows[j]["TimeColumnq"]?.ToString() + "','" + uSrno + "');";
                                    dem++;
                                }
                            }
                        }
                        catch
                        {
                        }
                    }
                    string cnnInsert34 = GetConnectionString("34");
                    bool a = SqlCnn.ExecuteNonQuery(sql_insert, cnnInsert34);
                    string cnnInsert4 = GetConnectionString("4");
                    bool b = SqlCnn.ExecuteNonQuery(sql_insert, cnnInsert4);
                    if (a & b)
                        //if (a )
                    {
                        AddMailRow(may, "Thành công máy BB ", prodat1, dem);
                    }
                    else
                    {
                        AddMailRow(may, "Insert Thất bại máy BB ", prodat1, dem);
                    }
                }
                else
                {
                    AddMailRow(may, "Máy BB hôm qua không chạy ", prodat1, dem);
                }
            }
            catch (Exception)
            {
                AddMailRow(may, "Văng Exception ex BB , máy có thể tắt ", DateTime.Now.AddDays(-1.0).ToString(TodayFormat), 0);
            }
        }

        static int? GetPtypeFromRubno7(DataTable table, string rubno7)
        {
            foreach (DataRow row in table.Rows)
            {
                if (row["rubno_7"].ToString() == rubno7)
                {
                    return Convert.ToInt32(row["ptype"]);
                }
            }
            return null;
        }

        public static string mail_Chuyenlieulot(DataTable dt, string day)
        {
            if (dt.Rows.Count > 0)
            {
                StringBuilder HtmlBuilder = new StringBuilder();
                HtmlBuilder.Append("<html >");
                HtmlBuilder.Append("<head>");
                HtmlBuilder.Append("</head>");
                HtmlBuilder.Append("<body>");
                HtmlBuilder.Append("<p style = 'font-size:16px;font-weight:bold;'>Mail Báo cáo chuyển liệu LOT NEW " + day + " </p>");
                HtmlBuilder.Append("<table border = '1' cellpadding='1' cellspacing='1'  style='font-family:Time new roman'>");
                HtmlBuilder.Append("<tr style='background-color:skyblue'>");
                foreach (DataColumn mycomlumn in dt.Columns)
                {
                    HtmlBuilder.Append("<td width='18%' >");
                    HtmlBuilder.Append("<center>");
                    HtmlBuilder.Append("<b>");
                    HtmlBuilder.Append(mycomlumn.ColumnName);
                    HtmlBuilder.Append("</b>");
                    HtmlBuilder.Append("</center>");
                    HtmlBuilder.Append("</td>");
                }
                HtmlBuilder.Append("</tr>");

                foreach (DataRow myRow in dt.Rows)
                {
                    HtmlBuilder.Append("<tr >");
                    foreach (DataColumn myColumn in dt.Columns)
                    {
                        HtmlBuilder.Append("<td >");
                        HtmlBuilder.Append("<center>");
                        HtmlBuilder.Append(myRow[myColumn.ColumnName].ToString());
                        HtmlBuilder.Append("</center>");
                        HtmlBuilder.Append("</td>");
                    }
                    HtmlBuilder.Append("</tr>");
                }
                HtmlBuilder.Append("</table>");
                HtmlBuilder.Append("</body>");
                HtmlBuilder.Append("</html>");
                string htmltext = HtmlBuilder.ToString();

                Send_Email_Without_Attachment(htmltext, "Mail Báo cáo chuyển liệu LotNEW ! " + DateTime.Now.ToString("dd/MM/yyyy") + " " + DateTime.Now.ToString("HH:mm:ss"));
            }
            return "";
        }
        public static bool Send_Email_Without_Attachment(string body, string Subject)
        {
            bool bResult = true;
            try
            {
                MailMessage em = new MailMessage();
                em.From = new System.Net.Mail.MailAddress("kenda_kv@kenda.com.tw");

                em.To.Add("tnamit@kenda.com.tw");
                em.CC.Add("loiit@kenda.com.tw");

                em.Subject = Subject;
                em.Body = body;
                em.IsBodyHtml = true;
                //System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("kmail.kenda.com.tw");
                //smtp.EnableSsl = false;
                //smtp.Credentials = new System.Net.NetworkCredential();
                //smtp.Send(em);


                //mailnew
                SmtpClient smtp = new SmtpClient("kmail.kenda.com.tw", 25);
                smtp.EnableSsl = false;

                // QUAN TRỌNG:
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = null;

                
                smtp.Send(em);





                em.Dispose();
                smtp.Dispose();

            }
            catch (Exception)
            {
                bResult = false;
            }
            return bResult;
        }
        public void insertBBMESMOI(string may)
        {
            try
            {

                {

                    int dem = 0;
                    string prodat1 = DateTime.Now.AddDays(-1).ToString(TodayFormat);

                    //prodat1 = "20260325";
                    string SQLVlevel = "SELECT ptype,rubno_7 FROM [InTem].[dbo].[rubnod_Ptype]";
                    string cnn186 = GetConnectionString("186");
                    DataTable DTcheckvlevel = SqlCnn.ExecuteQuery(SQLVlevel, cnn186);

                    string sql_insert = "";
                    string cnn33 = GetConnectionString("33");
                    string sql = "SELECT mesid,barcode,partno,weight,slipno,prodat,class,machno,indat,intime,usrno,some_sx  FROM [erp].[dbo].[prdebe] where LEFT([mesid], 1) NOT IN ('V', 'E','R','') and machno ='" + may + "' and prodat  like '" + prodat1 + "%'  order by indat , intime ";

                   // string sql = "SELECT mesid,barcode,partno,weight,slipno,prodat,class,machno,indat,intime,usrno,some_sx  FROM [erp].[dbo].[prdebe] where machno ='" + may + "'  and " +
                        //"barcode IN ('RB26410069','RB26410068','RB26410067') and mesid !='RL' and len(mesid) >14  order by indat , intime ";
                    DataTable dt = SqlCnn.ExecuteQuery(sql, cnn33);
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            try
                            {
                                string mes = dt.Rows[i]["mesid"].ToString().Trim();
                                string barcode = dt.Rows[i]["barcode"].ToString().Trim();
                                string pArtno = dt.Rows[i]["partno"].ToString().Trim();
                                string wEight = dt.Rows[i]["weight"].ToString().Trim();
                                string sLipno = dt.Rows[i]["slipno"].ToString().Trim();
                                string pRodat = dt.Rows[i]["prodat"].ToString().Trim();
                                string cLass = dt.Rows[i]["class"].ToString().Trim();
                                string mAchno = dt.Rows[i]["machno"].ToString().Trim();
                                string iNdat = dt.Rows[i]["indat"].ToString().Trim();
                                string iNtime = dt.Rows[i]["intime"].ToString().Trim();
                                string uSrno = dt.Rows[i]["usrno"].ToString().Trim();
                                string sMesx = dt.Rows[i]["some_sx"].ToString().Trim();
                                int? ptype = GetPtypeFromRubno7(DTcheckvlevel, pArtno);
                                string level = GetRubberLevel(barcode, ptype);


                                DataTable dtDatangguyenlieu = new DataTable();


                                string sqlPlan_id = "SELECT  [Plan_ID] FROM [mfns].[dbo].[Ppt_GroupLot]  where MesPlanID='" + mes + "'";
                                string cnnMay = GetConnectionString(may.Trim().Substring(may.Length - 2));
                                DataTable dtPlanid = SqlCnn.ExecuteQuery(sqlPlan_id, cnnMay);
                                if (dtPlanid.Rows.Count > 0)
                                {
                                    string plan_id = dtPlanid.Rows[0][0].ToString().Trim();

                                 
                                   
                                    var listMe = string.IsNullOrEmpty(sMesx)
                                                ? new List<string>()
                                                : sMesx.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)
                                                       .Select(x => x.Trim())
                                                       .ToList();

                                    string sqlIn = BuildSqlIn(listMe);
                                    var listMeSau = new List<string>(listMe);

                                    for (int c = 0;c <= 2; c++)
                                    {
                                       
                                        string query = BuildMaterialQuery(plan_id, sqlIn, true);

                                        dtDatangguyenlieu = SqlCnn.ExecuteQuery(query, cnnMay);

                                        if (dtDatangguyenlieu.Rows.Count > 0)
                                            break;


                                       
                                        listMe = listMe
                                            .Select(x => (int.Parse(x) - 1).ToString())
                                            .ToList();

                                        sqlIn = BuildSqlIn(listMe);
                                    }
                                    if (dtDatangguyenlieu.Rows.Count == 0)
                                    {
                                     

                                        listMe = new List<string>(listMeSau);

                                        sqlIn = BuildSqlIn(listMe);

                                        for (int c = 0; c <= 2; c++)
                                        {
                                            

                                            string query = BuildMaterialQuery(plan_id, sqlIn, true);

                                            dtDatangguyenlieu = SqlCnn.ExecuteQuery(query, cnnMay);

                                            if (dtDatangguyenlieu.Rows.Count > 0)
                                                break;

                                            listMe = listMe
                                                .Select(x => (int.Parse(x) + 1).ToString())
                                                .ToList();

                                            sqlIn = BuildSqlIn(listMe);
                                        }

                                    }

                                  
                                    var listBarcode = listMe
                                                .Select(x => plan_id + int.Parse(x).ToString("000"))
                                                .ToList();
                                    string sqlBarcodeIn = BuildSqlIn(listBarcode);


                                    string dataDau = "  SELECT a.[barcode],a.[mater_code], a.[equip_code], a.[set_weight], CONVERT(nvarchar(20),a.[weigh_time],120) as weigh_time ,a.[error_allow], a.[weigh_type]," +
                                        " b.mater_name  FROM[mfns].[dbo].[ppt_weigh] a, [mfns].[dbo].[pmt_material] b where barcode in " + sqlBarcodeIn + " " +
                                        " and weigh_type='油料' and a.mater_code = b.mater_code  order by weigh_time asc ";
                                    DataTable dtDatangguyenlieuDau = SqlCnn.ExecuteQuery(dataDau, cnnMay);
                                    if (dtDatangguyenlieuDau.Rows.Count != 0)
                                    {
                                        string s_fromday = dtDatangguyenlieuDau.Rows[0]["weigh_time"].ToString().Trim();
                                        DateTime dat_point = DateTime.Parse(s_fromday);
                                        DateTime dat_check = DateTime.Parse(s_fromday);
                                        TimeSpan ts_check = new TimeSpan(06, 30, 00);
                                        dat_check = dat_check.Date + ts_check;
                                        if (dat_point <= dat_check)
                                        {
                                            dat_check = dat_check.AddDays(-1);
                                        }
                                     
                                        List<string> coalCodes = new List<string>();

                                        for (int _icount = 0; _icount < dtDatangguyenlieuDau.Rows.Count; _icount++)
                                        {
                                            if (dtDatangguyenlieuDau.Rows[_icount]["weigh_type"].ToString().Trim() == "油料")
                                            {
                                                string materCode = dtDatangguyenlieuDau.Rows[_icount]["mater_code"].ToString().Trim();

                                                if (!coalCodes.Contains(materCode))
                                                {
                                                    coalCodes.Add(materCode);
                                                }
                                            }
                                        }
                                      
                                        string s_coal_barcode = @"
                                 SELECT Mater_Code, Mater_Barcode, SaveTime
                                 FROM (
                                     SELECT Mater_Code, Mater_Barcode, SaveTime,
                                            ROW_NUMBER() OVER (PARTITION BY Mater_Code ORDER BY SaveTime DESC) AS rn
                                     FROM [mfns].[dbo].[Ppt_Oil]
                                     WHERE SaveTime <= '" + s_fromday + @"'
                                       AND Mater_Type = '0'
                                       AND Mater_Code IN " + BuildSqlIn(coalCodes) + @"
                                 ) t
                                 WHERE rn = 1";


                                        DataTable dt_coal_barcode = SqlCnn.ExecuteQuery(s_coal_barcode, cnnMay);


                                     

                                        var barcodeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                        foreach (DataRow r in dt_coal_barcode.Rows)
                                        {
                                            string mcode = r["Mater_Code"]?.ToString().Trim() ?? "";
                                            string mbar = r["Mater_Barcode"]?.ToString().Trim() ?? "";
                                            if (!string.IsNullOrEmpty(mcode) && !string.IsNullOrEmpty(mbar))
                                            {
                                                // nếu trùng khóa, giữ cái đầu vì đã là mới nhất (rn=1)
                                                if (!barcodeMap.ContainsKey(mcode)) barcodeMap[mcode] = mbar;
                                            }
                                        }


                                        DataRow dr;
                                        int i_max = dtDatangguyenlieuDau.Rows.Count;
                                        for (int i_count = 0; i_count < i_max; i_count++)
                                        {
                                            dr = dtDatangguyenlieu.NewRow();

                                         

                                            if (barcodeMap.TryGetValue(dtDatangguyenlieuDau.Rows[i_count]["mater_code"].ToString().Trim(), out var matchedBarcode) && !string.IsNullOrEmpty(matchedBarcode))
                                            {
                                                if (matchedBarcode.Length >= 13 && matchedBarcode.StartsWith(dtDatangguyenlieuDau.Rows[i_count]["mater_code"].ToString().Trim(), StringComparison.OrdinalIgnoreCase))
                                                {
                                                    dr[0] = matchedBarcode;
                                                    dr[1] = dtDatangguyenlieuDau.Rows[i_count]["mater_code"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    dr[0] = OilNoData;
                                                    dr[1] = dtDatangguyenlieuDau.Rows[i_count]["mater_code"].ToString().Trim();
                                                }
                                            }
                                            else
                                            {
                                                dr[0] = OilNoData;
                                                dr[1] = dtDatangguyenlieuDau.Rows[i_count]["mater_code"].ToString().Trim();
                                            }





                                            dr[2] = dtDatangguyenlieuDau.Rows[i_count]["set_weight"].ToString();
                                            dr[3] = dtDatangguyenlieuDau.Rows[i_count]["weigh_time"].ToString().Trim().Substring(0, 10);
                                            dr[4] = dtDatangguyenlieuDau.Rows[i_count]["weigh_time"].ToString().Trim().Substring(11);
                                            dr[5] = int.Parse(dtDatangguyenlieuDau.Rows[i_count]["barcode"].ToString().Trim().Substring(dtDatangguyenlieuDau.Rows[i_count]["barcode"].ToString().Trim().Length - 2, 2));
                                            dr[6] = dtDatangguyenlieuDau.Rows[i_count]["barcode"].ToString();
                                            dtDatangguyenlieu.Rows.Add(dr);
                                        }
                                    }

                                }

                                if (dtDatangguyenlieu.Rows.Count > 0)
                                {
                                    for (int j = 0; j < dtDatangguyenlieu.Rows.Count; j++)
                                    {

                                        sql_insert += "insert into LOT_PCR values('" + level + "','" + barcode + "','" + pArtno + "','" + wEight + "'," +
                                            " '" + sLipno + "','" + pRodat + "','" + cLass + "','" + mAchno + "','" + iNdat + "','" + iNtime + "'," +
                                            " '" + uSrno + "','','','','','','','','','','','','','','','','" + dtDatangguyenlieu.Rows[j]["Mater_Barcode"] + "'," +
                                            " '" + dtDatangguyenlieu.Rows[j]["Mater_Code"] + "','" + dtDatangguyenlieu.Rows[j]["real_weight"] + "','" + dtDatangguyenlieu.Rows[j]["DateColumn"] + "'," +
                                            " '" + dtDatangguyenlieu.Rows[j]["TimeColumnq"] + "','" + uSrno + "');";
                                        dem = dem + 1;
                                    }
                                }
                            }
                            catch { continue; }

                        }
                        string cnnInsert34 = GetConnectionString("34");
                        bool a = SqlCnn.ExecuteNonQuery(sql_insert, cnnInsert34);
                        string cnnInsert4 = GetConnectionString("4");
                        bool b = SqlCnn.ExecuteNonQuery(sql_insert, cnnInsert4);
                        if (a && b)
                            //if (a )
                        {
                            AddMailRow(may, "(Mesmoi)Thành công máy BB ", prodat1, dem);
                        }
                        else
                        {
                            AddMailRow(may, "(Mesmoi)Insert Thất bại máy BB ", prodat1, dem);
                        }
                    }
                    else
                    {
                        AddMailRow(may, "(Mesmoi)Máy BB hôm qua không chạy ", prodat1, dem);
                    }

                }
            }
            catch (Exception)
            {
                AddMailRow(may, "(Mesmoi)Văng Exception ex BB , máy có thể tắt ", DateTime.Now.AddDays(-1).ToString(TodayFormat), 0);
            }


        }
    }
}
