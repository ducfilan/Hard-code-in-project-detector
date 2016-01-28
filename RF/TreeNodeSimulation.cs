using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RF
{
    public class TreeNodeSimulation
    {
        public TreeNodeSimulation(string name, string tag, int imageIndex)
        {
            Name = name;
            Tag = tag;
            ImageIndex = imageIndex;

            Nodes = new List<TreeNodeSimulation>();
        }

        public string Name { get; set; }
        public string Tag { get; set; }
        public int ImageIndex { get; set; }

        public List<TreeNodeSimulation> Nodes { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}
