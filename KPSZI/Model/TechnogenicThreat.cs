using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Техногенная угроза БИ
    /// </summary>
    public class TechnogenicThreat
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int TechnogenicThreatId { get; set; }

        /// <summary>
        /// Идентификатор типа ТУ.01 etc.
        /// </summary>
        public string Identificator { get; set; }

        /// <summary>
        /// Описание техногенной угрозы
        /// </summary>
        public string Description { get; set; }
    }
}
