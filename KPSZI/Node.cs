using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    class Node : PictureBox
    {
        NodeType nodeType;
        Stage stage;
        internal int centerX, centerY;

        public Node(Point location, NodeType nodeType, StageTopology stage)
        {
            centerX = location.X;
            centerY = location.Y;
            this.nodeType = nodeType;
            this.stage = stage;
            Size = new Size(64, 64);
            SizeMode = PictureBoxSizeMode.Zoom;
            BorderStyle = BorderStyle.None;
            Location = new Point(location.X - Size.Width / 2, location.Y - Size.Height / 2);
            Image = Properties.Resources.pc;

            switch (nodeType)
            {
                case (NodeType.Pc):
                    Image = Properties.Resources.pc;
                    Size = Properties.Resources.pc.Size;
                    break;
                case (NodeType.Server):
                    Image = Properties.Resources.server;
                    Size = Properties.Resources.server.Size;
                    break;
                case (NodeType.Switch):
                    Image = Properties.Resources._switch;
                    Size = Properties.Resources._switch.Size;
                    break;
                case (NodeType.Router):
                    Image = Properties.Resources.router;
                    Size = Properties.Resources.router.Size;
                    break;
                case (NodeType.Cloud):
                    Image = Properties.Resources.cloud;
                    Size = Properties.Resources.cloud.Size;
                    break;
            }

            MouseClick += new MouseEventHandler(stage.node_MouseClick);
            MouseMove += new MouseEventHandler(stage.node_MouseMove);
            MouseDown += new MouseEventHandler(stage.node_MouseDown);
            MouseUp += new MouseEventHandler(stage.node_MouseUp);
            LocationChanged += new EventHandler(stage.node_LocationChanged);
        }
    }
}
