using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace KPSZI
{
    enum Mode
    {
        Node,
        Link,
        Move,
        Delete
    }

    enum NodeType
    {
        Pc,
        Server,
        Switch,
        Router,
        Cloud
    }
    
    class StageTopology : Stage
    {
        protected override ImageList imageListForTabPage { get; set; }
        
        List<Node> nodes;
        List<Link> links;
        internal Point firstLoc, secondLoc;
        bool ySet = false;
        internal Mode currentMode;
        internal NodeType nodeType;
        Point DownPoint;
        bool IsDragMode;
        Node firstNode, secondNode;
        int mouseX = 0, mouseY = 0;
        Link selectedLink;

        public StageTopology(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {
            
            nodes = new List<Node>();
            links = new List<Link>();

            firstLoc = new Point();
            secondLoc = new Point();

            mf.pTopology.MouseMove += new MouseEventHandler(panel_MouseMove);
            mf.pTopology.MouseUp += new MouseEventHandler(panel_MouseUp);
            mf.pTopology.MouseClick += new MouseEventHandler(panel_MouseClick);
            mf.pTopology.Paint += new PaintEventHandler(panel_Paint);
            mf.tsbMove.Click += new EventHandler(tsbMove_Click);
            mf.tsbLink.Click += new EventHandler(tsbLink_Click);
            mf.tsbDelete.Click += new EventHandler(tsbDelete_Click);
            mf.tsbPc.Click += new EventHandler(tsbPc_Click);
            mf.tsbServer.Click += new EventHandler(tsbServer_Click);
            mf.tsbRouter.Click += new EventHandler(tsbRouter_Click);
            mf.tsbSwitch.Click += new EventHandler(tsbSwitch_Click);
            mf.KeyDown += new KeyEventHandler(Form_KeyDown);

            setMode(Mode.Move);
        }

        public override void enterTabPage()
        {

        }

        public override void saveChanges()
        {

        }

        protected override void initTabPage()
        {

        }

        private void setNodeType(NodeType nt)
        {
            nodeType = nt;
            mf.tsbPc.Checked = false;
            mf.tsbServer.Checked = false;
            mf.tsbRouter.Checked = false;
            mf.tsbSwitch.Checked = false;

            switch (nodeType)
            {
                case (NodeType.Pc):
                    mf.tsbPc.Checked = true;
                    break;
                case (NodeType.Server):
                    mf.tsbServer.Checked = true;
                    break;
                case (NodeType.Router):
                    mf.tsbRouter.Checked = true;
                    break;
                case (NodeType.Switch):
                    mf.tsbSwitch.Checked = true;
                    break;
                case (NodeType.Cloud):
                    mf.tsbSwitch.Checked = true;
                    break;
            }
        }

        private void setMode(Mode mode)
        {
            currentMode = mode;
            mf.tsbLink.Checked = false;
            mf.tsbMove.Checked = false;
            mf.tsbDelete.Checked = false;

            switch (mode)
            {
                case Mode.Link:
                    mf.tsbLink.Checked = true;
                    break;
                case Mode.Move:
                    mf.tsbMove.Checked = true;
                    break;
                case Mode.Delete:
                    mf.tsbDelete.Checked = true;
                    break;
            }
        }

        private void panel_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void panel_MouseMove(object sender, MouseEventArgs e)
        {
            if (firstNode != null && ySet == false)
                secondLoc = new Point(e.X, e.Y);

            mouseX = e.X;
            mouseY = e.Y;

            mf.pTopology.Invalidate();
        }

        public void drawLine(Point p1, Point p2, PaintEventArgs e)
        {
            Pen pen = new Pen(Color.Black);
            e.Graphics.DrawLine(pen, p1, p2);
        }

        public double areaTriangle(Point p1, Point p2, Point p3)
        {
            double AB = Math.Sqrt(Math.Pow(p2.X - p1.X, 2) + Math.Pow(p2.Y - p1.Y, 2));
            double BC = Math.Sqrt(Math.Pow(p3.X - p2.X, 2) + Math.Pow(p3.Y - p2.Y, 2));
            double CA = Math.Sqrt(Math.Pow(p1.X - p3.X, 2) + Math.Pow(p1.Y - p3.Y, 2));
            double p = (AB + BC + CA) / 2;

            return Math.Sqrt(p * (p - AB) * (p - BC) * (p - CA));
        }

        private void panel_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            if (firstNode != null && ySet == false)
                using (var p = new Pen(Color.Black))
                    e.Graphics.DrawLine(p, new Point(firstNode.Location.X + firstNode.Size.Width / 2, firstNode.Location.Y + firstNode.Size.Height / 2), secondLoc);


            foreach (Link link in links)
            {
                Point p1 = new Point();
                Point p2 = new Point();
                Point p3 = new Point();
                Point p4 = new Point();

                int lineH = 10;
                double linkLength = Math.Sqrt(Math.Pow(Math.Abs(link.firstNode.centerX - link.secondNode.centerX), 2) +
                    Math.Pow(Math.Abs(link.firstNode.centerY - link.secondNode.centerY), 2));
                double linkAreaRect = linkLength * lineH;
                double linkAngle = Math.Asin(Math.Abs(link.firstNode.centerY - link.secondNode.centerY) / linkLength);


                if ((link.firstNode.centerY - link.secondNode.centerY) * (link.firstNode.centerX - link.secondNode.centerX) >= 0)
                {
                    p1 = new Point(link.firstNode.centerX + (int)(Math.Sin(linkAngle) * lineH / 2), link.firstNode.centerY - (int)(Math.Cos(linkAngle) * lineH / 2));
                    p2 = new Point(link.secondNode.centerX + (int)(Math.Sin(linkAngle) * lineH / 2), link.secondNode.centerY - (int)(Math.Cos(linkAngle) * lineH / 2));
                    p3 = new Point(link.firstNode.centerX - (int)(Math.Sin(linkAngle) * lineH / 2), link.firstNode.centerY + (int)(Math.Cos(linkAngle) * lineH / 2));
                    p4 = new Point(link.secondNode.centerX - (int)(Math.Sin(linkAngle) * lineH / 2), link.secondNode.centerY + (int)(Math.Cos(linkAngle) * lineH / 2));
                }
                else
                {
                    p1 = new Point(link.firstNode.centerX + (int)(Math.Sin(linkAngle) * lineH / 2), link.firstNode.centerY + (int)(Math.Cos(linkAngle) * lineH / 2));
                    p2 = new Point(link.secondNode.centerX + (int)(Math.Sin(linkAngle) * lineH / 2), link.secondNode.centerY + (int)(Math.Cos(linkAngle) * lineH / 2));
                    p3 = new Point(link.firstNode.centerX - (int)(Math.Sin(linkAngle) * lineH / 2), link.firstNode.centerY - (int)(Math.Cos(linkAngle) * lineH / 2));
                    p4 = new Point(link.secondNode.centerX - (int)(Math.Sin(linkAngle) * lineH / 2), link.secondNode.centerY - (int)(Math.Cos(linkAngle) * lineH / 2));
                }

                double triangleSumm = areaTriangle(p1, p2, new Point(mouseX, mouseY)) + areaTriangle(p3, p4, new Point(mouseX, mouseY)) +
                    areaTriangle(p1, p3, new Point(mouseX, mouseY)) + areaTriangle(p2, p4, new Point(mouseX, mouseY));

                Pen p = new Pen(Color.Black);

                if (triangleSumm <= linkAreaRect && currentMode == Mode.Delete)
                {
                    p = new Pen(Color.Red);
                    selectedLink = link;
                }


                e.Graphics.DrawLine(p,
                    new Point(link.firstNode.Location.X + link.firstNode.Size.Width / 2, link.firstNode.Location.Y + link.firstNode.Size.Height / 2),
                    new Point(link.secondNode.Location.X + link.secondNode.Size.Width / 2, link.secondNode.Location.Y + link.secondNode.Size.Height / 2));
            }
        }

        private void panel_MouseClick(object sender, MouseEventArgs e)
        {
            if (currentMode == Mode.Node)
            {
                Node newNode = new Node(new Point(e.X, e.Y), nodeType, this);

                mf.pTopology.Controls.Add(newNode);
                nodes.Add(newNode);
                //setMode(Mode.Move);
            }

            if (currentMode == Mode.Delete)
            {
                if (selectedLink != null)
                    links.Remove(selectedLink);
            }

            mf.pTopology.Invalidate();
        }

        internal void node_MouseClick(object sender, MouseEventArgs e)
        {
            if (currentMode == Mode.Link)
            {
                if (firstNode == null)
                {
                    firstNode = (Node)sender;
                    ySet = false;
                    return;
                }
                else
                {
                    secondNode = (Node)sender;
                    if (secondNode.GetHashCode() == firstNode.GetHashCode())
                    {
                        firstNode = null;
                        secondNode = null;
                        mf.pTopology.Invalidate();
                        MessageBox.Show("Циклическая связь недопустима", "Ошибка создания связи", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    foreach (Link link in links)
                        if ((link.firstNode.GetHashCode() == firstNode.GetHashCode() && link.secondNode.GetHashCode() == secondNode.GetHashCode()) ||
                            (link.firstNode.GetHashCode() == secondNode.GetHashCode() && link.secondNode.GetHashCode() == firstNode.GetHashCode()))
                        {
                            firstNode = null;
                            secondNode = null;
                            mf.pTopology.Invalidate();
                            MessageBox.Show("Данная связь уже существует", "Ошибка создания связи", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                    links.Add(new Link(firstNode, secondNode));
                    firstNode = null;
                    ySet = true;
                    //setMode(Mode.Move);
                    return;
                }
            }

            if (currentMode == Mode.Delete)
            {
                mf.pTopology.Controls.Remove((Node)sender);
                foreach (Link link in links)
                {
                    if (link.firstNode.Equals((Node)sender) || link.secondNode.Equals((Node)sender))
                    {
                        links.Remove(link);
                        node_MouseClick(sender, e);
                        return;
                    }
                }
            }
        }

        internal void node_MouseMove(object sender, MouseEventArgs e)
        {
            Node currentNode = (Node)sender;

            if (currentMode == Mode.Link)
            {
                if (firstNode != null && ySet == false)
                    secondLoc = new Point(e.X + currentNode.Location.X, e.Y + currentNode.Location.Y);
                mf.pTopology.Invalidate();
                return;
            }

            if (currentMode == Mode.Move)
            {
                if (e.Button == MouseButtons.Left)
                {
                    if (IsDragMode && e.Button == MouseButtons.Left)
                        currentNode.Location = new Point(currentNode.Location.X + e.Location.X - DownPoint.X, currentNode.Location.Y + e.Location.Y - DownPoint.Y);
                }
            }
        }

        internal void node_MouseDown(object sender, MouseEventArgs e)
        {
            if (currentMode == Mode.Move)
            {
                DownPoint = e.Location;
                IsDragMode = true;
            }
        }

        internal void node_MouseUp(object sender, MouseEventArgs e)
        {
            if (currentMode == Mode.Move)
            {
                IsDragMode = false;
            }
        }

        internal void node_LocationChanged(object sender, System.EventArgs e)
        {
            mf.pTopology.Invalidate();
            Node n = ((Node)sender);
            n.centerX = n.Location.X + n.Width / 2;
            n.centerY = n.Location.Y + n.Height / 2;
        }

        private void tsbMove_Click(object sender, EventArgs e)
        {
            setMode(Mode.Move);
        }

        private void tsbPc_Click(object sender, EventArgs e)
        {
            setMode(Mode.Node);
            setNodeType(NodeType.Pc);
        }

        private void tsbServer_Click(object sender, EventArgs e)
        {
            setMode(Mode.Node);
            setNodeType(NodeType.Server);
        }

        private void tsbSwitch_Click(object sender, EventArgs e)
        {
            setMode(Mode.Node);
            setNodeType(NodeType.Switch);
        }

        private void tsbRouter_Click(object sender, EventArgs e)
        {
            setMode(Mode.Node);
            setNodeType(NodeType.Router);
        }

        private void tsbCloud_Click(object sender, EventArgs e)
        {
            setMode(Mode.Node);
            setNodeType(NodeType.Cloud);
        }

        private void tsbLink_Click(object sender, EventArgs e)
        {
            setMode(Mode.Link);
        }

        private void tsbDelete_Click(object sender, EventArgs e)
        {
            setMode(Mode.Delete);
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (mf.tabControl.SelectedTab == stageTab && e.KeyCode == Keys.Escape && ySet == false && firstNode != null)
            {
                firstNode = null;
                mf.pTopology.Invalidate();
            }
        }
    }
}
