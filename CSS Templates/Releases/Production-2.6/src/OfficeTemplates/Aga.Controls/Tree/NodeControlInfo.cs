using System.Drawing;
using Aga.Controls.Tree.NodeControls;

namespace Aga.Controls.Tree
{
    public struct NodeControlInfo
    {
        public static readonly NodeControlInfo Empty = new NodeControlInfo(null, Rectangle.Empty, null);
        private readonly Rectangle _bounds;

        private readonly NodeControl _control;
        private readonly TreeNodeAdv _node;

        public NodeControlInfo(NodeControl control, Rectangle bounds, TreeNodeAdv node)
        {
            _control = control;
            _bounds = bounds;
            _node = node;
        }

        public NodeControl Control
        {
            get { return _control; }
        }

        public Rectangle Bounds
        {
            get { return _bounds; }
        }

        public TreeNodeAdv Node
        {
            get { return _node; }
        }
    }
}