using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI.Model
{
    /// <summary>
    /// Класс-сущность для справочника 
    /// </summary>
    public class ThreatSource
    {
        /// <summary>
        /// Id для записи
        /// </summary>
        public int ThreatSourceId { get; set ; }
        /// <summary>
        /// Внутренний ли нарушитель? true - да, false - нет.
        /// </summary>
        public bool InternalIntruder { get; set; }
        /// <summary>
        /// Потенциал нарушителя. 0 - низкий, 1 - средний, 2 - высокий, 3 - нарушитель вообще не подразумевается.
        /// </summary>
        public int Potencial { get; set; }
        /// <summary>
        /// Коллекция всех угроз, для которых характерен данный источник угроз
        /// </summary>
        public virtual ICollection<Threat> Threats { get; set; }

        
        public ThreatSource()
        {
            Threats = new List<Threat>();
        }

        /// <summary>
        /// Метод парсит текстовое описание источника угрозы и приводит к объекту класса ThreatSource
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static List<ThreatSource> Parse(string str, List<ThreatSource> tss)
        {
            // Список распарсенных источников угроз
            List<ThreatSource> listOfTS = new List<ThreatSource>();

            if (str == "" || str == null)
            {
                ThreatSource tsnull = tss.Where(t => t.Potencial == 3).FirstOrDefault();
                listOfTS.Add(tsnull);
                return listOfTS;
            }

            string[] arrOfTextTS = str.Split(',');
            
            // Если массив пустой, то добавляем в него всю строку:
            if (arrOfTextTS == null)
            {
                arrOfTextTS = new string[1] { str };
            }

            foreach (string s in arrOfTextTS)
            {
                ThreatSource ts = new ThreatSource();

                if (s.ToLower().Contains("внутренний"))
                {
                    if (s.ToLower().Contains("низким"))
                        ts = tss.Where(t => t.InternalIntruder == true && t.Potencial == 0).FirstOrDefault();
                    else if (s.ToLower().Contains("средним"))
                        ts = tss.Where(t => t.InternalIntruder == true && t.Potencial == 1).FirstOrDefault();
                    else if (s.ToLower().Contains("высоким"))
                        ts = tss.Where(t => t.InternalIntruder == true && t.Potencial == 2).FirstOrDefault();
                }
                else if (s.ToLower().Contains("внешний"))
                {
                    if (s.ToLower().Contains("низким"))
                        ts = tss.Where(t => t.InternalIntruder == false && t.Potencial == 0).FirstOrDefault();
                    else if (s.ToLower().Contains("средним"))
                        ts = tss.Where(t => t.InternalIntruder == false && t.Potencial == 1).FirstOrDefault();
                    else if (s.ToLower().Contains("высоким"))
                        ts = tss.Where(t => t.InternalIntruder == false && t.Potencial == 2).FirstOrDefault();
                }
                else
                {
                    MessageBox.Show("Произошла ошибка парсинга поля 'Источник угроз'.\nПриступай к дебаггингу", "Ахтунг!",  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                listOfTS.Add(ts);
            }

            return listOfTS;
        }

    }
}
