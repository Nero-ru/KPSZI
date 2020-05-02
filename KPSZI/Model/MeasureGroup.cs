using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Группа мер в приказах 27/21, например, ИАФ, УПД, ЗСВ, ЗТС 
    /// </summary>
    public class MeasureGroup
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int MeasureGroupId { get; set;}
        
        /// <summary>
        /// Полное название группы мер
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// Сокращенное названия группы мер
        /// </summary>
        public string ShortName { get; set; }

        /// <summary>
        /// Описание группы мер
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Навигационное поле - коллекция мер для ИСПДН в группе
        /// </summary>
        public virtual ICollection<ISPDNMeasure> ISPDNMeasures { get; set; }

        /// <summary>
        /// Навигационное поле - коллекция мер для ГИС в группе
        /// </summary>
        public virtual ICollection<GISMeasure> GISMeasures { get; set; }

        /// <summary>
        /// Конструктор для создания коллекций
        /// </summary>
        public MeasureGroup ()
        {
            this.ISPDNMeasures = new List<ISPDNMeasure>();
            this.GISMeasures = new List<GISMeasure>();
        }
    }
}
