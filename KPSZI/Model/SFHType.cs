using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Класс справочник типов структурно-функциональных характеристик (По архитектуре, по подкл. и другим сетям, etc.)
    /// </summary>
    public class SFHType
    {
        /// <summary>
        /// Первичный ключ - Id типа СФХ
        /// </summary>
        public int SFHTypeId { get; set; }
        /// <summary>
        /// Наименование типа СФХ
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Можно ли выбрать несколько СФХ?
        /// </summary>
        public Boolean MultipleChoice { get; set; }
        /// <summary>
        /// Навигационное поле для выгрузки СФХ из таблицы
        /// </summary>
        public virtual ICollection<SFH> SFHs { get; set; }
        /// <summary>
        /// Конструктор для создания списка
        /// </summary>
        public SFHType()
        {
            this.SFHs = new List<SFH>();
        }

        public override bool Equals(object obj)
        {
            var sfhtype = (SFHType)obj;

            if (sfhtype.Name == Name)
                return true;

            return false;
        }

        public override int GetHashCode()
        {
            return Name.Count();
        }
    }
}
