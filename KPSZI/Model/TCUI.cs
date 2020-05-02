using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Техический канал утечки информации
    /// </summary>
    public class TCUI
    {
        /// <summary>
        /// Первичный ключ Id для таблицы
        /// </summary>
        public int TCUIId { get; set; }
        
        /// <summary>
        /// Наименование ТКУИ
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Навигационное поле для хранения типа
        /// </summary>
        public virtual TCUIType TCUIType { get; set;}
        
        /// <summary>
        /// Навигационное поле для хранения угроз утечки по ТК
        /// </summary>
        public virtual ICollection< TCUIThreat > TCUIThreats { get; set; }

        /// <summary>
        /// Конструктор для создания коллекции
        /// </summary>
        public TCUI()
        {
            this.TCUIThreats = new List<TCUIThreat>();
        }
    }
}
