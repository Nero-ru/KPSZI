using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Мера защиты от техногенной УБИ
    /// </summary>
    public class TechnogenicMeasure
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int TechnogenicMeasureId { get; set; }

        /// <summary>
        /// Описание меры
        /// </summary>
        public string Description { get; set; }
    }
}
