using System;

namespace Aga.Controls.Tree.NodeControls
{
    public class NodeEventArgs : EventArgs
    {
        private readonly TreeNodeAdv _node;

        public NodeEventArgs(TreeNodeAdv node)
        {
            _node = node;
        }

        public TreeNodeAdv Node
        {
            get { return _node; }
        }
    }
}