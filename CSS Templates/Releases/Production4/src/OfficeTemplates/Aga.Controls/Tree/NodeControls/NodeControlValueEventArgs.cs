namespace Aga.Controls.Tree.NodeControls
{
    public class NodeControlValueEventArgs : NodeEventArgs
    {
        public NodeControlValueEventArgs(TreeNodeAdv node)
            : base(node)
        {
        }

        public object Value { get; set; }
    }
}