using System;
using System.Collections;

namespace Aga.Controls.Tree
{
    /// <summary>
    ///     Converts IEnumerable interface to ITreeModel.
    ///     Allows to display a plain list in the TreeView
    /// </summary>
    public class TreeListAdapter : ITreeModel
    {
        private readonly IEnumerable _list;

        public TreeListAdapter(IEnumerable list)
        {
            _list = list;
        }

        #region ITreeModel Members

        public IEnumerable GetChildren(TreePath treePath)
        {
            if (treePath.IsEmpty())
                return _list;
            else
                return null;
        }

        public bool IsLeaf(TreePath treePath)
        {
            return true;
        }

        public event EventHandler<TreeModelEventArgs> NodesChanged;

        public event EventHandler<TreePathEventArgs> StructureChanged;

        public event EventHandler<TreeModelEventArgs> NodesInserted;

        public event EventHandler<TreeModelEventArgs> NodesRemoved;

        public void OnNodesChanged(TreeModelEventArgs args)
        {
            if (NodesChanged != null)
                NodesChanged(this, args);
        }

        public void OnStructureChanged(TreePathEventArgs args)
        {
            if (StructureChanged != null)
                StructureChanged(this, args);
        }

        public void OnNodeInserted(TreeModelEventArgs args)
        {
            if (NodesInserted != null)
                NodesInserted(this, args);
        }

        public void OnNodeRemoved(TreeModelEventArgs args)
        {
            if (NodesRemoved != null)
                NodesRemoved(this, args);
        }

        #endregion
    }
}