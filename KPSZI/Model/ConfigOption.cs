using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Параметр настройки
    /// </summary>
    public class ConfigOption
    {
        /// <summary>
        /// Ключ
        /// </summary>
        public int ConfigOptionId { get; set; }
        /// <summary>
        /// Мера, к которой относится параметр
        /// </summary>
        public virtual GISMeasure GISMeasure { get; set; }

        /// <summary>
        /// Название, определение параметра
        /// </summary>
        public string Name { get; set; }
    
        /// <summary>
        /// Описание, уточнение параметра
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Минимальный класс защищенности ИС для необходимости параметра
        /// 0 - все; 1, 2, 3 - соответственно
        /// </summary>
        public string DefenceClass { get; set; }
    }
}
