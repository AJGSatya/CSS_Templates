using System;
using System.Collections;

namespace Aga.Controls.Tree
{
    public abstract class TreeModelBase : ITreeModel
    {
        public abstract IEnumerable GetChildren(TreePath treePath);
        public abstract bool IsLeaf(TreePath treePath);


        public event EventHandler<TreeModelEventArgs> NodesChanged;

        public event EventHandler<TreePathEventArgs> StructureChanged;

        public event EventHandler<TreeModelEventArgs> NodesInserted;

        public event EventHandler<TreeModelEventArgs> NodesRemoved;

        protected void OnNodesChanged(TreeModelEventArgs args)
        {
            if (NodesChanged != null)
                NodesChanged(this, args);
        }

        protected void OnStructureChanged(TreePathEventArgs args)
        {
            if (StructureChanged != null)
                StructureChanged(this, args);
        }

        protected void OnNodesInserted(TreeModelEventArgs args)
        {
            if (NodesInserted != null)
                NodesInserted(this, args);
        }

        protected void OnNodesRemoved(TreeModelEventArgs args)
        {
            if (NodesRemoved != null)
                NodesRemoved(this, args);
        }

        public virtual void Refresh()
        {
            OnStructureChanged(new TreePathEventArgs(TreePath.Empty));
        }
    }
}