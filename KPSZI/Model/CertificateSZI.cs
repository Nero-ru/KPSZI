using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.ComponentModel.DataAnnotations.Schema;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.Globalization;

namespace KPSZI.Model
{
    /// <summary>
    /// Сертификаты СЗИ из ресстра ФСТЭК
    /// </summary
    public class CertificateSZI
    {
        /// <summary>
        /// ID угрозы
        /// </summary>
        [Key]
        public int? CertificateId { get; set; }

        /// <summary>
        /// Номер сертификата по списку ФСТЭКа
        /// </summary>
        public string CertificateNumber { get; set; }

        /// <summary>
        /// Наименование СЗИ
        /// </summary>
        public string NameSZI { get; set; }

        /// <summary>
        /// Срок действия сертификата
        /// </summary>
        public string Validity { get; set; }

        /// <summary>
        /// Дата окончания технической поддержки 
        /// </summary>
        public string ValidityTechnicalSupport { get; set; }

        /// <summary>
        /// Метод выстаскивает текстовое описание сертификатов, парсит, приводит к типу CertificateSZI и возвращает массив объектов CertificateSZI.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static List<CertificateSZI> GetThreatsFromXlsx(FileInfo file, KPSZIContext db)
        {
            List<List<string>> listOfXlslxRows = File_Selected_New(file);

            List<CertificateSZI> listOfAllThreatsFromFile = new List<CertificateSZI>();
            for (int i = 0; i < listOfXlslxRows[0].Count; i++)
            {
                CertificateSZI сSZI = new CertificateSZI();
                сSZI.CertificateNumber = listOfXlslxRows[0][i].ToString();
                сSZI.Validity = listOfXlslxRows[2][i].ToString();
                сSZI.NameSZI = listOfXlslxRows[3][i];
                сSZI.ValidityTechnicalSupport = listOfXlslxRows[10][i].ToString();
                listOfAllThreatsFromFile.Add(сSZI);
            }

            return listOfAllThreatsFromFile;
        }

        private static List<List<string>> File_Selected_New(FileInfo _file)
        {
            DataSet ds = new DataSet();
            string connectionString = GetConnectionString(_file);

            using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                conn.Open();
                System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    ds.Tables.Add(dt);
                }
                // В list_table хранится информация по строкам, начиная с [0]
                List<List<string>> list_table = new List<List<string>>();
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    List<string> list_row = new List<string>();
                    for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
                        list_row.Add(ds.Tables[0].Rows[i].ItemArray[j].ToString());
                    list_table.Add(list_row);
                }

                return list_table;
            }
        }
        private static string GetConnectionString(FileInfo _file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            if (_file.Extension == ".xlsx")
            {
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
                props["Data Source"] = _file.FullName;
            }
            else if (_file.Extension == ".xls")
            {
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
                props["Data Source"] = "C:\\MyExcel.xls";
            }
            else throw new Exception("Неизвестное расширение файла Реестра ФСТЭК СЗИ \"_reestr_sszi.ods\"!");

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

    }
}
