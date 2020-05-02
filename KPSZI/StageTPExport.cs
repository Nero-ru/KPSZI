using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPSZI
{
    class StageTPExport:Stage
    {
        protected override ImageList imageListForTabPage { get; set; }

        public StageTPExport(TabPage stageTab, TreeNode stageNode, MainForm mainForm, InformationSystem IS)
            : base(stageTab, stageNode, mainForm, IS)
        {

        }

        protected override void initTabPage()
        {

        }

        public override void saveChanges()
        {

        }

        public override void enterTabPage()
        {

        }
    }
}
