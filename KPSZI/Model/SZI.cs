using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    /// <summary>
    /// Средства защиты информации
    /// </summary>
    public class SZI
    {
        /// <summary>
        /// Первичный ключ
        /// </summary>
        public int SZIId { get; set; }
        /// <summary>
        /// Класс защиты, который обеспечивает данное СЗИ, если равно нулю, 
        /// то требования к классам защиты для данного вида не устанавливаются.
        /// </summary>
        public string DefenceClass { get; set; }
        /// <summary>
        /// Наименование СЗИ
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Номер сертификата, моэет содержать символы
        /// </summary>
        public string Certificate { get; set; }
        /// <summary>
        /// Дата окончания срока действия сертификата  
        /// </summary>
        public DateTime DateOfEnd { get; set; }
        /// <summary>
        /// Уровень контроля отсутствия НДВ
        /// </summary>
        public int NDVControlLevel { get; set; }
        /// <summary>
        /// Класс СВТ
        /// </summary>
        public string SVTClass { get; set; }
        /// <summary>
        /// Сертифицирован ли по ТУ, если да, то ставится крестик
        /// </summary>
        public string TU { get; set; }
        /// <summary>
        /// Системные требования к частоте ЦПУ 
        /// </summary>
        public int CPUFrequencyRequirements { get; set; }
        /// <summary>
        /// Системные требования к количеству ядер ЦПУ
        /// </summary>
        public int CPUCoresRequirements { get; set; }
        /// <summary>
        /// Системные требования к оперативной памяти
        /// </summary>
        public int MemoryRequirements { get; set; }
        /// <summary>
        /// Системные требования к свободному месту на диске
        /// </summary>
        public int DiscSpaceRequirements { get; set; }
        /// <summary>
        /// Коллекция видов СЗИ, к которым относится данное
        /// </summary>
        public virtual ICollection<SZISort> SZISorts { get; set; }
        /// <summary>
        /// Список мер ЗИ в ГИС, реализуемых данным СЗИ - навигационное поле
        /// </summary>
        public virtual ICollection<GISMeasure> GISMeasures { get; set; }
        /// <summary>
        /// Список мер ЗИ в ИСПДН, реализуемых данным СЗИ - навигационное поле
        /// </summary>
       // public virtual ICollection<ISPDNMeasure> ISPDNMeasures { get; set; }

        public SZI()
        {
            this.SZISorts= new List<SZISort>();
            this.GISMeasures = new List<GISMeasure>();
            SVTClass = "";
            TU = "";
            CPUFrequencyRequirements = 0;
            CPUCoresRequirements = 0;
            MemoryRequirements = 0;
            DiscSpaceRequirements = 0;
            //this.ISPDNMeasures = new List<ISPDNMeasure>();
        }

        public override bool Equals(object obj)
        {
            SZI szi;
            try
            {
                szi = (SZI)obj;
            }
            catch
            {
                return false;
            }
            if (this.Name == szi.Name)
                return true;

            return false;
        }

        public override int GetHashCode()
        {
            return Name.Length + DateOfEnd.GetHashCode();
        }
    }
}
