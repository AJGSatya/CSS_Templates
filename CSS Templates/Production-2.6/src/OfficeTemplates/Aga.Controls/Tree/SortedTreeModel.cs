using System.Collections;

namespace Aga.Controls.Tree
{
    public class SortedTreeModel : TreeModelBase
    {
        private readonly ITreeModel _innerModel;

        private IComparer _comparer;

        public SortedTreeModel(ITreeModel innerModel)
        {
            _innerModel = innerModel;
            _innerModel.NodesChanged += _innerModel_NodesChanged;
            _innerModel.NodesInserted += _innerModel_NodesInserted;
            _innerModel.NodesRemoved += _innerModel_NodesRemoved;
            _innerModel.StructureChanged += _innerModel_StructureChanged;
        }

        public ITreeModel InnerModel
        {
            get { return _innerModel; }
        }

        public IComparer Comparer
        {
            get { return _comparer; }
            set
            {
                _comparer = value;
                OnStructureChanged(new TreePathEventArgs(TreePath.Empty));
            }
        }

        private void _innerModel_StructureChanged(object sender, TreePathEventArgs e)
        {
            OnStructureChanged(e);
        }

        private void _innerModel_NodesRemoved(object sender, TreeModelEventArgs e)
        {
            OnStructureChanged(new TreePathEventArgs(e.Path));
        }

        private void _innerModel_NodesInserted(object sender, TreeModelEventArgs e)
        {
            OnStructureChanged(new TreePathEventArgs(e.Path));
        }

        private void _innerModel_NodesChanged(object sender, TreeModelEventArgs e)
        {
            OnStructureChanged(new TreePathEventArgs(e.Path));
        }

        public override IEnumerable GetChildren(TreePath treePath)
        {
            if (Comparer != null)
            {
                var list = new ArrayList();
                IEnumerable res = InnerModel.GetChildren(treePath);
                if (res != null)
                {
                    foreach (object obj in res)
                        list.Add(obj);
                    list.Sort(Comparer);
                    return list;
                }
                else
                    return null;
            }
            else
                return InnerModel.GetChildren(treePath);
        }

        public override bool IsLeaf(TreePath treePath)
        {
            return InnerModel.IsLeaf(treePath);
        }
    }
}