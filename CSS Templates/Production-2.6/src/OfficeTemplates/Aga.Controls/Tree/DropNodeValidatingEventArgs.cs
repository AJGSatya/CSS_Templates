using System;
using System.Drawing;

namespace Aga.Controls.Tree
{
    public class DropNodeValidatingEventArgs : EventArgs
    {
        private readonly Point _point;

        public DropNodeValidatingEventArgs(Point point, TreeNodeAdv node)
        {
            _point = point;
            Node = node;
        }

        public Point Point
        {
            get { return _point; }
        }

        public TreeNodeAdv Node { get; set; }
    }
}