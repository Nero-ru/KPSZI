using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    public class InfoType
    {
        //Справочник видов информации
        public int InfoTypeId { get; set; }
        /// <summary>
        /// Вид информации
        /// </summary>
        public string TypeName { get; set; }
    }
}
