using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPSZI.Model
{
    public class ImplementWay
    {
        public int ImplementWayId { get; set; }
        /// <summary>
        /// Номер способа реализации
        /// </summary>
        public int WayNumber { get; set; }
        /// <summary>
        /// Способ реализации УБИ
        /// </summary>
        public string WayName { get; set; }
        /// <summary>
        /// Коллекция всех угроз, для которых характерен данный способ реализации УБИ
        /// </summary>
        public virtual ICollection<Threat> Threats { get; set; }

        public ImplementWay()
        {
            Threats = new List<Threat>();
        }
    }
}
