using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;

namespace task_delta_mes_pc_inventory
{
    public class MES_SERVER_INVENTORY
    {
        public string ID { get; set; }

        public string IPV4 { get; set; }

        public string NAME { get; set; }

        public string DESCRIPTION { get; set; }

        public string DIVISION { get; set; }

        public string OWNER { get; set; }

        public string CDATE { get; set; }

        public string OS { get; set; }

        public string CATEGORY { get; set; }

        public string OS_NAME { get; set; }

        public string SERVER_TYPE_NAME { get; set; }

        public string OWNER_NAME { get; set; }

        public string DIVISION_NAME { get; set; }

        public string CREATE_BY { get; set; }
    }

    public class MES_SERVER_INVENTORY_INPUT_NAME
    {
        public string DIVISION_NAME { get; set; }
        public string SERVER_TYPE_NAME { get; set; }
    }

    public class MES_SERVER_INVENTORY_GROUP
    {
        public double y { get; set; }

        public string name { get; set; }

    }

    public static class ServerInventoryAction
    {
        #region *** Property
        private static string dbcon = "Provider=sqloledb;Data Source=THBPOCIMDB; User Id=MESDB;Password=MES12345;";

        public static List<MES_SERVER_INVENTORY> Get()
        {
            return Get(new MES_SERVER_INVENTORY());
        }
        #endregion

        #region *** Main
        public static List<MES_SERVER_INVENTORY> Get(MES_SERVER_INVENTORY data)
        {
            List<MES_SERVER_INVENTORY> retData = new List<MES_SERVER_INVENTORY>();
            string NAME = "";
            if (data != null)
            {
                if (data.NAME != "")
                {
                    NAME = $"AND NAME = '{data.NAME}'";
                }
            }

            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string SQL = $@"SELECT SVINV.*
                        ,OS_MT.OS_NAME 
                        ,SERVER_TYPE_MT.SERVER_TYPE_NAME
                        ,DIVISION_MT.DIVISION_NAME
                        ,OWNER_MT.OWNER_NAME
                        FROM MESPRDDB.dbo.MES_SERVER_INVENTORY SVINV                       
                        LEFT JOIN (SELECT ID OS_ID,NAME OS_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'OS') OS_MT
                        ON SVINV.OS  = OS_MT.OS_ID
                        LEFT JOIN (SELECT ID SERVER_TYPE_ID,NAME SERVER_TYPE_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'SERVER_TYPE') SERVER_TYPE_MT
                        ON SVINV.CATEGORY = SERVER_TYPE_MT.SERVER_TYPE_ID
                        LEFT JOIN (SELECT ID DIVISION_ID,NAME DIVISION_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'DIVISION') DIVISION_MT
                        ON SVINV.DIVISION = DIVISION_MT.DIVISION_ID
                        LEFT JOIN (SELECT ID OWNER_ID,NAME OWNER_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'SERVER_OWNER') OWNER_MT
                        ON SVINV.OWNER = OWNER_MT.OWNER_ID
                        WHERE 1=1  {NAME}   
                        ORDER BY CDATE DESC";
                DB.ExecQuery(SQL, out DataTable DT, out string Status);
                for (int idx = 0; idx < DT.Rows.Count; idx++)
                {
                    retData.Add(new MES_SERVER_INVENTORY
                    {
                        ID = MyDataTable.GetCell(DT, "ID", "", idx),
                        NAME = MyDataTable.GetCell(DT, "NAME", "", idx),
                        OWNER = MyDataTable.GetCell(DT, "OWNER", "", idx),
                        CDATE = MyDataTable.GetCell(DT, "CDATE", "", idx),
                        OS = MyDataTable.GetCell(DT, "OS", "", idx),
                        DESCRIPTION = MyDataTable.GetCell(DT, "DESCRIPTION", "", idx),
                        DIVISION = MyDataTable.GetCell(DT, "DIVISION", "", idx),
                        IPV4 = MyDataTable.GetCell(DT, "IPV4", "", idx),
                        CATEGORY = MyDataTable.GetCell(DT, "CATEGORY", "", idx),
                        OS_NAME = MyDataTable.GetCell(DT, "OS_NAME", "", idx),
                        SERVER_TYPE_NAME = MyDataTable.GetCell(DT, "SERVER_TYPE_NAME", "", idx),
                        OWNER_NAME = MyDataTable.GetCell(DT, "OWNER_NAME", "", idx),
                        DIVISION_NAME = MyDataTable.GetCell(DT, "DIVISION_NAME", "", idx),
                    });
                }
                return retData;
            }
        }

        public static List<MES_SERVER_INVENTORY> Get(MES_SERVER_INVENTORY_INPUT_NAME data)
        {
            List<MES_SERVER_INVENTORY> retData = new List<MES_SERVER_INVENTORY>();
            string DIVISION_NAME = ""; string SERVER_TYPE_NAME = "";
            if (data != null)
            {
                if (data.SERVER_TYPE_NAME != null)
                {
                    SERVER_TYPE_NAME = $"AND SERVER_TYPE_NAME = '{data.SERVER_TYPE_NAME}'";
                }

                if (data.DIVISION_NAME != null)
                {
                    DIVISION_NAME = $"AND DIVISION_NAME = '{data.DIVISION_NAME}'";
                }
            }

            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string SQL = $@"SELECT SVINV.*
                        ,OS_MT.OS_NAME 
                        ,SERVER_TYPE_MT.SERVER_TYPE_NAME
                        ,DIVISION_MT.DIVISION_NAME
                        ,OWNER_MT.OWNER_NAME
                        FROM MESPRDDB.dbo.MES_SERVER_INVENTORY SVINV                       
                        LEFT JOIN (SELECT ID OS_ID,NAME OS_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'OS') OS_MT
                        ON SVINV.OS  = OS_MT.OS_ID
                        LEFT JOIN (SELECT ID SERVER_TYPE_ID,NAME SERVER_TYPE_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'SERVER_TYPE') SERVER_TYPE_MT
                        ON SVINV.CATEGORY = SERVER_TYPE_MT.SERVER_TYPE_ID
                        LEFT JOIN (SELECT ID DIVISION_ID,NAME DIVISION_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'DIVISION') DIVISION_MT
                        ON SVINV.DIVISION = DIVISION_MT.DIVISION_ID
                        LEFT JOIN (SELECT ID OWNER_ID,NAME OWNER_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'SERVER_OWNER') OWNER_MT
                        ON SVINV.OWNER = OWNER_MT.OWNER_ID
                        WHERE 1=1 {DIVISION_NAME} {SERVER_TYPE_NAME}  
                        ORDER BY CDATE DESC";
                DB.ExecQuery(SQL, out DataTable DT, out string Status);
                for (int idx = 0; idx < DT.Rows.Count; idx++)
                {
                    retData.Add(new MES_SERVER_INVENTORY
                    {
                        ID = MyDataTable.GetCell(DT, "ID", "", idx),
                        NAME = MyDataTable.GetCell(DT, "NAME", "", idx),
                        OWNER = MyDataTable.GetCell(DT, "OWNER", "", idx),
                        CDATE = MyDataTable.GetCell(DT, "CDATE", "", idx),
                        OS = MyDataTable.GetCell(DT, "OS", "", idx),
                        DESCRIPTION = MyDataTable.GetCell(DT, "DESCRIPTION", "", idx),
                        DIVISION = MyDataTable.GetCell(DT, "DIVISION", "", idx),
                        IPV4 = MyDataTable.GetCell(DT, "IPV4", "", idx),
                        CATEGORY = MyDataTable.GetCell(DT, "CATEGORY", "", idx),
                        OS_NAME = MyDataTable.GetCell(DT, "OS_NAME", "", idx),
                        SERVER_TYPE_NAME = MyDataTable.GetCell(DT, "SERVER_TYPE_NAME", "", idx),
                        OWNER_NAME = MyDataTable.GetCell(DT, "OWNER_NAME", "", idx),
                        DIVISION_NAME = MyDataTable.GetCell(DT, "DIVISION_NAME", "", idx),
                    });
                }
                return retData;
            }
        }

        public static bool Insert(MES_SERVER_INVENTORY data, out string msgSQL)
        {
            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string ID = "null";
                if (data.ID != "")
                {
                    ID = $"'{data.ID}'";
                }

                string IPV4 = "null";
                if (data.IPV4 != "")
                {
                    IPV4 = $"'{data.IPV4}'";
                }

                string CDATE = "null";
                if (data.CDATE != "")
                {
                    CDATE = $"convert(datetime, '{data.CDATE}', 120)";
                }

                string OWNER = "null";
                if (data.OWNER != "")
                {
                    OWNER = $"'{data.OWNER}'";
                }

                string DIVISION = "null";
                if (data.DIVISION != "")
                {
                    DIVISION = $"'{data.DIVISION}'";
                }

                string DESCRIPTION = "null";
                if (data.DESCRIPTION != "")
                {
                    DESCRIPTION = $"'{data.DESCRIPTION}'";
                }

                string NAME = "null";
                if (data.NAME != "")
                {
                    NAME = $"'{data.NAME}'";
                }

                string OS = "null";
                if (data.OS != "")
                {
                    OS = $"'{data.OS}'";
                }

                string CATEGORY = "null";
                if (data.CATEGORY != "")
                {
                    CATEGORY = $"'{data.CATEGORY}'";
                }

                string CREATE_BY = "null";
                if (data.CREATE_BY != null)
                {
                    CREATE_BY = $"'{data.CREATE_BY}'";
                }

                string sql = $@"INSERT INTO MESPRDDB.dbo.MES_SERVER_INVENTORY
                        (ID,IPV4,CDATE,OWNER,DIVISION,DESCRIPTION,NAME,OS,CATEGORY,CREATE_BY)VALUES
                        ({ID},{IPV4},{CDATE},{OWNER},{DIVISION},{DESCRIPTION},{NAME},{OS},{CATEGORY},{CREATE_BY})";
                return DB.ExecNonQuery(sql, out msgSQL);
            }
        }

        public static bool Update(MES_SERVER_INVENTORY data, out string msgSQL)
        {
            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string ID = "null";
                if (data.ID != "")
                {
                    ID = $"'{data.ID}'";
                }

                string IPV4 = "null";
                if (data.IPV4 != "")
                {
                    IPV4 = $"'{data.IPV4}'";
                }

                string CDATE = "null";
                if (data.CDATE != "")
                {
                    CDATE = $"convert(datetime, '{data.CDATE}', 120)";
                }

                string OWNER = "null";
                if (data.OWNER != "")
                {
                    OWNER = $"'{data.OWNER}'";
                }

                string DIVISION = "null";
                if (data.DIVISION != "")
                {
                    DIVISION = $"'{data.DIVISION}'";
                }

                string DESCRIPTION = "null";
                if (data.DESCRIPTION != "")
                {
                    DESCRIPTION = $"'{data.DESCRIPTION}'";
                }

                string NAME = "null";
                if (data.NAME != "")
                {
                    NAME = $"'{data.NAME}'";
                }

                string OS = "null";
                if (data.OS != "")
                {
                    OS = $"'{data.OS}'";
                }

                string CATEGORY = "null";
                if (data.CATEGORY != "")
                {
                    CATEGORY = $"'{data.CATEGORY}'";
                }

                string CREATE_BY = "null";
                if (data.CREATE_BY != null)
                {
                    CREATE_BY = $"'{data.CREATE_BY}'";
                }

                string sql = $@"UPDATE MESPRDDB.dbo.MES_SERVER_INVENTORY SET
                            IPV4 = {IPV4}
                            ,CDATE = {CDATE}
                            ,OWNER = {OWNER}
                            ,DIVISION = {DIVISION}
                            ,DESCRIPTION = {DESCRIPTION}
                            ,NAME = {NAME}
                            ,OS = {OS}
                            ,CATEGORY = {CATEGORY}
                            ,CREATE_BY = {CREATE_BY}
                            WHERE ID = {ID}";

                return DB.ExecNonQuery(sql, out msgSQL);
            }
        }

        public static bool Delete(string id, out string msgSQL)
        {
            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string ID = "null";
                if (id != "")
                {
                    ID = $"'{id}'";
                }

                string sql = $@"DELETE FROM MESPRDDB.dbo.MES_SERVER_INVENTORY WHERE
                       id = {ID}";

                return DB.ExecNonQuery(sql, out msgSQL);
            }
        }

        #endregion

        #region *** Dashboard

        public static double GetTotalServer(MES_SERVER_INVENTORY data)
        {
            string IPV4 = ""; string OWNER = ""; string DIVISION = ""; string DESCRIPTION = ""; string NAME = ""; string OS = ""; string CATEGORY = "";
            if (data != null)
            {
                if (data.IPV4 != null)
                {
                    IPV4 = $"AND IPV4 = '{data.IPV4}'";
                }


                if (data.OWNER != null)
                {
                    OWNER = $"AND OWNER = '{data.OWNER}'";
                }


                if (data.DIVISION != null)
                {
                    DIVISION = $"AND DIVISION = '{data.DIVISION}'";
                }


                if (data.DESCRIPTION != null)
                {
                    DESCRIPTION = $"AND DESCRIPTION = '{data.DESCRIPTION}'";
                }


                if (data.NAME != null)
                {
                    NAME = $"AND NAME = '{data.NAME}'";
                }


                if (data.OS != null)
                {
                    OS = $"AND OS = '{data.OS}'";
                }


                if (data.CATEGORY != null)
                {
                    CATEGORY = $"AND CATEGORY = '{data.CATEGORY}'";
                }
            }
            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string SQL = $@"SELECT COUNT(*) total_server FROM MESPRDDB.dbo.MES_SERVER_INVENTORY
                        WHERE 1=1 {IPV4} {OWNER} {DIVISION} {DESCRIPTION} {NAME} {OS} {CATEGORY}  
                        ";
                DB.ExecQuery(SQL, out DataTable DT, out string Status);
                return double.Parse(MyDataTable.GetCell(DT, "total_server", "0"));
            }
        }

        public static List<MES_SERVER_INVENTORY_GROUP> GetTotalServerGroupByDivision(MES_SERVER_INVENTORY data)
        {
            List<MES_SERVER_INVENTORY_GROUP> result = new List<MES_SERVER_INVENTORY_GROUP>();

            string IPV4 = ""; string OWNER = ""; string DIVISION = ""; string DESCRIPTION = ""; string NAME = ""; string OS = ""; string CATEGORY = "";
            if (data != null)
            {
                if (data.IPV4 != null)
                {
                    IPV4 = $"AND IPV4 = '{data.IPV4}'";
                }


                if (data.OWNER != null)
                {
                    OWNER = $"AND OWNER = '{data.OWNER}'";
                }


                if (data.DIVISION != null)
                {
                    DIVISION = $"AND DIVISION = '{data.DIVISION}'";
                }


                if (data.DESCRIPTION != null)
                {
                    DESCRIPTION = $"AND DESCRIPTION = '{data.DESCRIPTION}'";
                }


                if (data.NAME != null)
                {
                    NAME = $"AND NAME = '{data.NAME}'";
                }

                if (data.OS != null)
                {
                    OS = $"AND OS = '{data.OS}'";
                }


                if (data.CATEGORY != null)
                {
                    CATEGORY = $"AND CATEGORY = '{data.CATEGORY}'";
                }
            }

            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string SQL = $@"SELECT total_server, DIVISION_NAME group_name FROM
                        (
                        SELECT COUNT(*) total_server,DIVISION FROM MESPRDDB.dbo.MES_SERVER_INVENTORY
                        WHERE 1=1 {IPV4} {OWNER} {DIVISION} {DESCRIPTION} {NAME} {OS} {CATEGORY}
                        GROUP BY DIVISION
                        )  SVINV
                        LEFT JOIN (SELECT ID DIVISION_ID,NAME DIVISION_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'DIVISION') DIVISION_MT
                        ON SVINV.DIVISION = DIVISION_MT.DIVISION_ID
                        ";
                DB.ExecQuery(SQL, out DataTable DT, out string Status);

                for (int idx = 0; idx < DT.Rows.Count; idx++)
                {
                    double total_server = double.Parse(MyDataTable.GetCell(DT, "total_server", "0", idx));
                    string group_name = MyDataTable.GetCell(DT, "group_name", "", idx);
                    result.Add(new MES_SERVER_INVENTORY_GROUP { name = group_name, y = total_server });
                }

                return result;
            }
        }

        public static List<MES_SERVER_INVENTORY_GROUP> GetTotalServerGroupByCategory(MES_SERVER_INVENTORY data)
        {
            List<MES_SERVER_INVENTORY_GROUP> result = new List<MES_SERVER_INVENTORY_GROUP>();

            string IPV4 = ""; string OWNER = ""; string DIVISION = ""; string DESCRIPTION = ""; string NAME = "";
            string OS = ""; string CATEGORY = "";

            if (data != null)
            {
                if (data.IPV4 != null)
                {
                    IPV4 = $"AND IPV4 = '{data.IPV4}'";
                }

                if (data.OWNER != null)
                {
                    OWNER = $"AND OWNER = '{data.OWNER}'";
                }


                if (data.DIVISION != null)
                {
                    DIVISION = $"AND DIVISION = '{data.DIVISION}'";
                }


                if (data.DESCRIPTION != null)
                {
                    DESCRIPTION = $"AND DESCRIPTION = '{data.DESCRIPTION}'";
                }


                if (data.NAME != null)
                {
                    NAME = $"AND NAME = '{data.NAME}'";
                }

                if (data.OS != null)
                {
                    OS = $"AND OS = '{data.OS}'";
                }


                if (data.CATEGORY != null)
                {
                    CATEGORY = $"AND CATEGORY = '{data.CATEGORY}'";
                }
            }

            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string SQL = $@"SELECT total_server, SERVER_TYPE_NAME group_name FROM
                        (
                        SELECT COUNT(*) total_server,CATEGORY FROM MESPRDDB.dbo.MES_SERVER_INVENTORY
                        WHERE 1=1 {IPV4} {OWNER} {DIVISION} {DESCRIPTION} {NAME} {OS} {CATEGORY}
                        GROUP BY CATEGORY
                        )  SVINV
                        LEFT JOIN (SELECT ID SERVER_TYPE_ID,NAME SERVER_TYPE_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'SERVER_TYPE') SERVER_TYPE_MT
                        ON SVINV.CATEGORY = SERVER_TYPE_MT.SERVER_TYPE_ID
                        ";
                DB.ExecQuery(SQL, out DataTable DT, out string Status);
                for (int idx = 0; idx < DT.Rows.Count; idx++)
                {
                    double total_server = double.Parse(MyDataTable.GetCell(DT, "total_server", "0", idx));
                    string group_name = MyDataTable.GetCell(DT, "group_name", "Other", idx);
                    result.Add(new MES_SERVER_INVENTORY_GROUP { name = group_name, y = total_server });
                }

                return result;
            }
        }

        public static List<MES_SERVER_INVENTORY_GROUP> GetTotalServerGroupByOS(MES_SERVER_INVENTORY data)
        {
            List<MES_SERVER_INVENTORY_GROUP> result = new List<MES_SERVER_INVENTORY_GROUP>();

            string IPV4 = ""; string OWNER = ""; string DIVISION = ""; string DESCRIPTION = ""; string NAME = ""; string OS = ""; string CATEGORY = "";
            if (data != null)
            {
                if (data.IPV4 != null)
                {
                    IPV4 = $"AND IPV4 = '{data.IPV4}'";
                }


                if (data.OWNER != null)
                {
                    OWNER = $"AND OWNER = '{data.OWNER}'";
                }


                if (data.DIVISION != null)
                {
                    DIVISION = $"AND DIVISION = '{data.DIVISION}'";
                }


                if (data.DESCRIPTION != null)
                {
                    DESCRIPTION = $"AND DESCRIPTION = '{data.DESCRIPTION}'";
                }


                if (data.NAME != null)
                {
                    NAME = $"AND NAME = '{data.NAME}'";
                }

                if (data.OS != null)
                {
                    OS = $"AND OS = '{data.OS}'";
                }


                if (data.CATEGORY != null)
                {
                    CATEGORY = $"AND CATEGORY = '{data.CATEGORY}'";
                }
            }

            using (MyOleDb DB = new MyOleDb(dbcon))
            {
                string SQL = $@"SELECT total_server, OS_NAME group_name FROM
                        (
                        SELECT COUNT(*) total_server,OS FROM MESPRDDB.dbo.MES_SERVER_INVENTORY
                        WHERE 1=1 {IPV4} {OWNER} {DIVISION} {DESCRIPTION} {NAME} {OS} {CATEGORY}
                        GROUP BY OS
                        )  SVINV
                        LEFT JOIN (SELECT ID OS_ID,NAME OS_NAME FROM MESPRDDB.dbo.MES_KEY_NAME_MT WHERE TYPE = 'OS') OS_MT
                        ON SVINV.OS = OS_MT.OS_ID
                        ";

                DB.ExecQuery(SQL, out DataTable DT, out string Status);

                for (int idx = 0; idx < DT.Rows.Count; idx++)
                {
                    double total_server = double.Parse(MyDataTable.GetCell(DT, "total_server", "0", idx));
                    string group_name = MyDataTable.GetCell(DT, "group_name", "", idx);
                    result.Add(new MES_SERVER_INVENTORY_GROUP { name = group_name, y = total_server });
                }

                return result;
            }
        }
        #endregion
    }
}
