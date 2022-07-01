using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ExcelDataReader;
using System.IO;
using Aspose.Cells;

namespace task_delta_mes_pc_inventory
{
    internal class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Reading XLSB file in C# using Aspose.Cells API.");
            Console.WriteLine("----------------------------------------------");

            Workbook WB = new Workbook(@"D:\Delta\Project\delta_mes_server_inventory\file\DataExport.xlsb");
            Worksheet WS = WB.Worksheets[0];
            DataTable DT = WS.Cells.ExportDataTableAsString(0, 0, WS.Cells.MaxRow, WS.Cells.MaxColumn);

            DT = MyDataTable.GetTableBySelect(DT, $"Column2 is not null");
            string msgSQL = "";
            double Total = DT.Rows.Count - 1;
            for (int idx = 1; idx < DT.Rows.Count - 1; idx++)
            {
                double progress = ((double.Parse((idx + 1).ToString("0")) / Total) * 100);

                Console.Write($"\r>>>>> Progress : {idx + 1}/{Total}, {progress.ToString("0.00")}%");

                string IPV4 = MyDataTable.GetCell(DT, "Column8", null, idx);
                IPV4 = IPV4.Replace("'", "");
                IPV4 = IPV4.Replace("]", "");
                IPV4 = IPV4.Replace("[", "");
                IPV4 = IPV4.Replace("None", "");

                string CDATE = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string OWNER = MyDataTable.GetCell(DT, "Column4", "", idx);
                string DIVISION = MyDataTable.GetCell(DT, "Column5", "", idx);
                string DESCRIPTION = "";
                string NAME = MyDataTable.GetCell(DT, "Column2", "", idx);
                string OS = MyDataTable.GetCell(DT, "Column7", "", idx);
                string CATEGORY = "";
                string CREATE_BY = "RPA";

                MES_SERVER_INVENTORY data = new MES_SERVER_INVENTORY();

                data.IPV4 = IPV4;
                data.CDATE = CDATE;
                data.OWNER = OWNER;
                data.DIVISION = DIVISION;
                data.CREATE_BY = CREATE_BY;
                data.DESCRIPTION = DESCRIPTION;
                data.NAME = NAME;
                data.OS = OS;
                data.CATEGORY = CATEGORY;

                List<MES_SERVER_INVENTORY> result = ServerInventoryAction.Get(data);
                if (result.Count > 0)
                {
                    data.ID = result[0].ID;
                    ServerInventoryAction.Update(data, out msgSQL);
                    if (msgSQL != "")
                    {
                        Console.WriteLine($", {idx} : {data.NAME}, {msgSQL}");
                    }
                }
                else
                {
                    Console.WriteLine($", New PC {idx} : {data.NAME}");
                    data.ID = DateTime.Now.ToString("yyyyMMddHHmmss") + data.NAME;
                    ServerInventoryAction.Insert(data, out msgSQL);

                    if (msgSQL != "")
                    {
                        Console.WriteLine($", {idx} : {data.NAME}, {msgSQL}");
                    }
                }

            }

            Console.ReadKey();
        }
    }
}
