using System;

namespace Aga.Controls.Tree
{
    public class TreeViewAdvEventArgs : EventArgs
    {
        private readonly TreeNodeAdv _node;

        public TreeViewAdvEventArgs(TreeNodeAdv node)
        {
            _node = node;
        }

        public TreeNodeAdv Node
        {
            get { return _node; }
        }
    }
}