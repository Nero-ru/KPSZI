using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Меры защиты информации в ГИС
    /// </summary>
    public class GISMeasure
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int GISMeasureId { get; set; }

        /// <summary>
        /// Номер меры - 1,2,3,... В связке с полем MeasureGroupId образуется идентификатор меры, например ИАФ.1, ЗСВ.5 
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// Навигационное поле - внешний ключ на группу мер
        /// </summary>
        public virtual MeasureGroup MeasureGroup { get; set; }

        /// <summary>
        /// Описание меры
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Минимально необходимый класс защиты ГИС, для которого должна применяться данная мера. 
        /// Возможные варианты: 3,2,1,0. Где 0 - значит не требуется КЗ ГИС.
        /// К примеру, =3 => в 3,2,1 требуется применение данной меры
        /// </summary>
        public int MinimalRequirementDefenceClass { get; set; }

        /// <summary>
        /// Флаг, если ==true, то мера указана только в 21 приказе и только для ИСПДн.
        /// Если ==false, то мера либо только для ГИС, либо для ГИС+ИСПДн.
        /// Между приказами 21 и 17 нет ситуаций, при которых аналогичные меры 
        /// не использовались бы в ГИС, являясь актуальными для ИСПДн.
        /// 
        /// </summary>
        public bool isOnlyISPDn { get; set; }
        
        /// <summary>
        /// Навигационное поле - коллекция СЗИ, реализующих меры
        /// </summary>
        public virtual ICollection<SZISort> SZISorts { get; set; }

        /// <summary>
        /// Коллекция угроз, которые нейтрализует данная мера [навигационное поле]
        /// </summary>
        public virtual ICollection<Threat> Threats { get; set; }
        /// <summary>
        /// Навигационное поле соответствия СФХ мерам для адаптации набора мер
        /// </summary>
        public virtual ICollection<SFH> SFHs { get; set; }

        /// <summary>
        /// Коллекция необходимых для реализации меры параметров настройки
        /// </summary>
        public virtual ICollection<ConfigOption> ConfigOptions { get; set; }

        /// <summary>
        /// Конструктор для инициализации коллекции навигационного поля
        /// </summary>
        public GISMeasure()
        {
            Threats = new List<Threat>();
            this.isOnlyISPDn = false;
            this.SZISorts = new List<SZISort>();
            SFHs = new List<SFH>();
            this.ConfigOptions = new List<ConfigOption>();
        }

        public override bool Equals(object obj)
        {
            GISMeasure gm;
            try
            {
                gm = (GISMeasure)obj;
            }
            catch
            {
                return false;
            }

            if (this.Description == gm.Description)
                return true;

            return false;
        }
        public override int GetHashCode()
        {
            return GISMeasureId + Description.Count();
        }

        public override string ToString()
        {
            return MeasureGroup.ShortName + "." + Number + ". " + Description;
        }
    }
}
