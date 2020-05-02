using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Меры защиты информации в ИСПДн
    /// </summary>
    public class ISPDNMeasure
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int ISPDNMeasureId { get; set; }

        /// <summary>
        /// Номер меры - 1,2,3,... В связке с полем MeasureGroupId образуется идентификатор меры, например ИАФ.1, ЗСВ.5 
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// Внешний ключ на группу мер, к примеру ИАФ, УПД
        /// </summary>
        public virtual MeasureGroup MeasureGroup { get; set; }

        /// <summary>
        /// Описание меры
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Минимально необходимый уровень защищенности ПДн, для которого должна применяться данная мера. К примеру, =3 => в 3,2,1 требуется применение данной меры
        /// </summary>
        public int MinimalRequirementDefenceLevel { get; set; }


        /// <summary>
        /// Навигационное поле - коллекция СЗИ, реализующих меры
        /// </summary>
        public virtual ICollection<SZI> SZIs { get; set; }

        /// <summary>
        /// Конструктор для инициализации коллекции навигационного поля
        /// </summary>
        public ISPDNMeasure()
        {
            this.SZIs = new List<SZI>();
        }
    }
}
