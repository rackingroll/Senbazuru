using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameFinder
{
    class HeaderNode
    {
        private int row;
        private int col;
        public int Row
        {
            get { return this.row; }
        }
        public int Col
        {
            get { return this.col; }
        }
        public List<HeaderNode> Children;
        public HeaderNode Parent;

        public HeaderNode()
        {
            this.row = -1;
            this.col = -1;
            this.Parent = null;
            this.Children = new List<HeaderNode>();
        }
        public HeaderNode(int row, int col)
        {
            this.row = row;
            this.col = col;
            this.Parent = null;
            this.Children = new List<HeaderNode>();
        }

        public void AddChild(HeaderNode child)
        {
            this.Children.Add(child);
        }

        public bool HasParent()
        {
            return this.Parent != null;
        }
    }
}
