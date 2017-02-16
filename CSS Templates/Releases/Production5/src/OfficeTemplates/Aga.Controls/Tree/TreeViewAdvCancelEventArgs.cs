namespace Aga.Controls.Tree
{
    public class TreeViewAdvCancelEventArgs : TreeViewAdvEventArgs
    {
        public TreeViewAdvCancelEventArgs(TreeNodeAdv node)
            : base(node)
        {
        }

        public bool Cancel { get; set; }
    }
}