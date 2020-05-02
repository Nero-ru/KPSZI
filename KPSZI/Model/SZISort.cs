using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Типы СЗИ - СОВ, СДЗ, etc.
    /// </summary>
    public class SZISort
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int SZISortId { get; set; }
        /// <summary>
        /// Название вида СЗИ
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Короткое название вида СЗИ
        /// </summary>
        public string ShortName { get; set; }
        /// <summary>
        /// Костыль для идентификации вида СЗИ в базе
        /// </summary>
        public int Number { get; set; }
        /// <summary>
        /// Список средств СЗИ данного вида
        /// </summary>
        public virtual ICollection<SZI> SZIs { get; set; }
        /// <summary>
        /// Меры, которые реализует данный вид СЗИ
        /// </summary>
        public virtual ICollection<GISMeasure> GISMeasures { get; set; }
        /// <summary>
        /// Конструктор для создания списка
        /// </summary>
        public SZISort()
        {
            this.SZIs = new List<SZI>();
        }
    }
}
