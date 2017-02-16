namespace Aga.Controls.Tree
{
    internal abstract class ColumnState : InputState
    {
        private readonly TreeColumn _column;

        public ColumnState(TreeViewAdv tree, TreeColumn column)
            : base(tree)
        {
            _column = column;
        }

        public TreeColumn Column
        {
            get { return _column; }
        }
    }
}