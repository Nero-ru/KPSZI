using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    class Link
    {
        internal Node firstNode, secondNode;

        public Link(Node first, Node second)
        {
            firstNode = first;
            secondNode = second;
        }
    }
}
