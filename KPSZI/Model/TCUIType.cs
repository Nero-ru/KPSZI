using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Классификация технических каналов утечки информации
    /// </summary>
    public class TCUIType
    {
        /// <summary>
        /// Первичный ключ - Id типа
        /// </summary>
        public int TCUITypeId { get; set; }

        /// <summary>
        /// Название типа
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Навигационное поле для угроз данного типа
        /// </summary>
        public virtual ICollection<TCUI> TCUIs { get; set; }

        /// <summary>
        /// Конструктор для создания списка
        /// </summary>
        public TCUIType()
        {
            this.TCUIs = new List<TCUI>();
        }
    }
}
