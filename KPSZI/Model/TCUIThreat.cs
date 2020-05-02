using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Класс угрозы утечки по техническим каналам
    /// </summary>
    public class TCUIThreat
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int TCUIThreatId { get; set; }

        /// <summary>
        /// Навигационное поле - внешний ключ на ТКУИ
        /// </summary>
        public virtual TCUI TCUI{ get; set; }

        /// <summary>
        /// Идентификатор угрозы, типа ТК.01 etc.
        /// </summary>
        public string Identificator { get; set; }

        /// <summary>
        /// Название угрозы
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// Описание угрозы
        /// </summary>
        public string Description { get; set; }
    }
}
