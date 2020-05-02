using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace KPSZI.Model
{

    /// <summary>
    ///  Класс самих СФ характеристик (Локальная, Распределенная, грид-системы, гетерогенная среда, etc.)
    /// </summary>
    public class SFH
    {
        /// <summary>
        /// Первичный ключ - Id СФХ
        /// </summary>
        public int SFHId { get; set; }

        /// <summary>
        /// Номер СФХ
        /// </summary>
        public int SFHNumber { get; set; }
        
        /// <summary>
        /// Навигационное поле - внешний ключ на тип СФХ
        /// </summary>
        public virtual SFHType SFHType { get; set; }

        /// <summary>
        ///  Наименование СФХ
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Уровень проектной защищенности: 2 - высокий, 1 - средний, 0 - низкий.
        /// </summary>
        public int ProjectSecurity { get; set; }
        /// <summary>
        /// Соответствие мер угрозам (уточнение адаптированного базового набора мер)
        /// </summary>
        public virtual ICollection<Threat> Threats { get; set; }
        /// <summary>
        /// Навигационное поле для соответствия мер защиты и СФХ (адаптация базового набора мер)
        /// </summary>
        public virtual ICollection<GISMeasure> GISMeasures { get; set; }

        public SFH()
        {
            GISMeasures = new List<GISMeasure>();
            Threats = new List<Threat>();
        }

        public override bool Equals(object obj)
        {
            var sfh = (SFH)obj;

            if (sfh.Name == Name)
                return true;

            return false;
        }

        public override int GetHashCode()
        {
            return Name.Count();
        }
    }
}
