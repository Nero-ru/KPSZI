using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.ComponentModel.DataAnnotations.Schema;
using System.Threading.Tasks;


namespace KPSZI.Model
{
    /// <summary>
    /// Класс-сущность угроза
    /// </summary>
    public class Threat
    {
        /// <summary>
        /// ID угрозы
        /// </summary>
        public int? ThreatId { get; set; }
        /// <summary>
        /// Номер угрозы по списку ФСТЭКа
        /// </summary>
        public int ThreatNumber { get; set; }
        /// <summary>
        /// Наименование УБИ
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Описание УБИ
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Объект воздействия
        /// </summary>
        public string ObjectOfInfluence { get; set; }
        /// <summary>
        /// Нарушение конфиденциальности
        /// </summary>
        public bool ConfidenceViolation { get; set; }
        /// <summary>
        /// Нарушение целостности
        /// </summary>
        public bool IntegrityViolation { get; set; }
        /// <summary>
        /// Нарушение доступности
        /// </summary>
        public bool AvailabilityViolation { get; set; }
        /// <summary>
        /// Дата добавления угрозы в БДУ
        /// </summary>
        public DateTime DateOfAdd { get; set; }
        /// <summary>
        /// Последняя дата изменения угрозы
        /// </summary>
        public DateTime DateOfChange { get; set; }
        /// <summary>
        /// Коллекция мер, нейтрализующих данную угрозу
        /// </summary>
        public virtual ICollection<GISMeasure> GISMeasures { get; set; }
        /// <summary>
        /// Коллекция всех источников угроз, которые характерны для данной угрозы [навигационное поле]
        /// </summary>
        public virtual ICollection<ThreatSource> ThreatSources { get; set; }
        public virtual ICollection<ImplementWay> ImplementWays { get; set; }        
        public virtual ICollection<Vulnerability> Vulnerabilities { get; set; }
        public virtual ICollection<SFH> SFHs { get; set; }
        [NotMapped]
        public string stringVuls { get; set; }
        [NotMapped]
        public string stringWays { get; set; }
        [NotMapped]
        public string stringSFHs { get; set; }
        [NotMapped]
        public string stringSources { get; set; }

        public Threat()
        {
            GISMeasures = new List<GISMeasure>();
            ThreatSources = new List<ThreatSource>();
            ImplementWays = new List<ImplementWay>();            
            Vulnerabilities = new List<Vulnerability>();
            SFHs = new List<SFH>();
        }

        public override bool Equals(object obj)
        {
            Threat thr;

            try
            {
                thr = (Threat)obj;
            }
            catch
            {
                return false;
            }

            if (thr.ThreatNumber == ThreatNumber && thr.Name == Name)
                return true;
            return false; 
        }

        public override int GetHashCode()
        {
            return ThreatNumber + Name.Length * 2;
        }

        public override string ToString()
        {
            return "УБИ." + ThreatNumber.ToString("000") + ". " + Name;
        }

        public void setStringImplementWays()
        {
            string s = "";
            foreach (ImplementWay iw in ImplementWays)
                s += iw.WayName + ";\n";
            if (s != "")
                s = s.Remove(s.Length - 1);
            stringWays = s;
        }

        public void setStringSources()
        {
            string s = "";
            int[] maxPotencial = new int[2] { -1, -1 };
            foreach (ThreatSource ts in ThreatSources)
            {
                int intern = ts.InternalIntruder ? 0 : 1;
                if (maxPotencial[intern] < ts.Potencial)
                    maxPotencial[intern] = ts.Potencial;
            }
            if (maxPotencial[0] != -1)
                s += "тип – внутренний, потенциал – " + (maxPotencial[0] == 0 ? "базовый (низкий)" :
                    (maxPotencial[0] == 1 ? "базовый (средний)" : "высокий"));
            if (maxPotencial[0] != -1 && maxPotencial[1] != -1)
                s += ";\n";
            if (maxPotencial[1] != -1)
                s += "тип – внешний, потенциал – " + (maxPotencial[1] == 0 ? "базовый (низкий)." :
                    (maxPotencial[1] == 1 ? "базовый (средний)." : "высокий."));
            if (maxPotencial[1] == 3)
                s = "Нарушитель не подразумевается.";
            stringSources = s;
        }

        public void setStringVulnerabilities()
        {
            string s = "";
            foreach (Vulnerability vul in Vulnerabilities)
                s += vul.VulnerabilityName + ";\n";
            if (s != "")
                s = s.Remove(s.Length - 1);
            stringVuls = s;
        }

        public void setStringSFHs()
        {
            string s = "";
            foreach (SFH sfh in SFHs)
                s += sfh.Name + ";\n";
            if (s != "")
                s = s.Remove(s.Length - 1);
            stringSFHs = s;
        }
        
        /// <summary>
        /// Метод выстаскивает текстовое описание угроз, парсит, приводит к типу Threat и возвращает массив объектов Threat.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static List<Threat> GetThreatsFromXlsx(FileInfo file, KPSZIContext db)
        {
            List<List<string>> listOfXlslxRows = File_Selected_New(file);
            List<ThreatSource> tss = db.ThreatSources.ToList();

            List<Threat> listOfAllThreatsFromFile = new List<Threat>();
            for (int i = 0; i < listOfXlslxRows[0].Count; i++)
            {
                Threat thr = new Threat();
                thr.ThreatNumber = Convert.ToInt32(listOfXlslxRows[0][i]);
                thr.Name = listOfXlslxRows[1][i];
                thr.Description = listOfXlslxRows[2][i];
                thr.ThreatSources = ThreatSource.Parse(listOfXlslxRows[3][i], tss); 
                thr.ObjectOfInfluence = listOfXlslxRows[4][i];
                thr.ConfidenceViolation = (listOfXlslxRows[5][i] == "1") ? true : false;
                thr.IntegrityViolation = (listOfXlslxRows[6][i] == "1") ? true : false;
                thr.AvailabilityViolation = (listOfXlslxRows[7][i] == "1") ? true : false;
                thr.DateOfAdd = DateTime.Parse(listOfXlslxRows[8][i]);
                thr.DateOfChange = DateTime.Parse(listOfXlslxRows[9][i]);                
                listOfAllThreatsFromFile.Add(thr);
            }

            return listOfAllThreatsFromFile;
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
            else throw new Exception("Неизвестное расширение файла списка угроз \"thrlist.xlsx\"!");

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
    }
}
