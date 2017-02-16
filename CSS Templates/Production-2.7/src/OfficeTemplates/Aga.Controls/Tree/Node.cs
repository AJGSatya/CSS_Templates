using System;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Windows.Forms;

namespace Aga.Controls.Tree
{
    public class Node
    {
        #region NodeCollection

        private class NodeCollection : Collection<Node>
        {
            private readonly Node _owner;

            public NodeCollection(Node owner)
            {
                _owner = owner;
            }

            protected override void ClearItems()
            {
                while (Count != 0)
                    RemoveAt(Count - 1);
            }

            protected override void InsertItem(int index, Node item)
            {
                if (item == null)
                    throw new ArgumentNullException("item");

                if (item.Parent != _owner)
                {
                    if (item.Parent != null)
                        item.Parent.Nodes.Remove(item);
                    item._parent = _owner;
                    item._index = index;
                    for (int i = index; i < Count; i++)
                        this[i]._index++;
                    base.InsertItem(index, item);

                    TreeModel model = _owner.FindModel();
                    if (model != null)
                        model.OnNodeInserted(_owner, index, item);
                }
            }

            protected override void RemoveItem(int index)
            {
                Node item = this[index];
                item._parent = null;
                item._index = -1;
                for (int i = index + 1; i < Count; i++)
                    this[i]._index--;
                base.RemoveItem(index);

                TreeModel model = _owner.FindModel();
                if (model != null)
                    model.OnNodeRemoved(_owner, index, item);
            }

            protected override void SetItem(int index, Node item)
            {
                if (item == null)
                    throw new ArgumentNullException("item");

                RemoveAt(index);
                InsertItem(index, item);
            }
        }

        #endregion

        #region Properties

        private readonly NodeCollection _nodes;
        private CheckState _checkState;
        private Image _image;
        private int _index = -1;

        private Node _parent;
        private string _text;
        internal TreeModel Model { get; set; }

        public Collection<Node> Nodes
        {
            get { return _nodes; }
        }

        public Node Parent
        {
            get { return _parent; }
            set
            {
                if (value != _parent)
                {
                    if (_parent != null)
                        _parent.Nodes.Remove(this);

                    if (value != null)
                        value.Nodes.Add(this);
                }
            }
        }

        public int Index
        {
            get { return _index; }
        }

        public Node PreviousNode
        {
            get
            {
                int index = Index;
                if (index > 0)
                    return _parent.Nodes[index - 1];
                else
                    return null;
            }
        }

        public Node NextNode
        {
            get
            {
                int index = Index;
                if (index >= 0 && index < _parent.Nodes.Count - 1)
                    return _parent.Nodes[index + 1];
                else
                    return null;
            }
        }

        public virtual string Text
        {
            get { return _text; }
            set
            {
                if (_text != value)
                {
                    _text = value;
                    NotifyModel();
                }
            }
        }

        public virtual CheckState CheckState
        {
            get { return _checkState; }
            set
            {
                if (_checkState != value)
                {
                    _checkState = value;
                    NotifyModel();
                }
            }
        }

        public Image Image
        {
            get { return _image; }
            set
            {
                if (_image != value)
                {
                    _image = value;
                    NotifyModel();
                }
            }
        }

        public object Tag { get; set; }

        public bool IsChecked
        {
            get { return CheckState != CheckState.Unchecked; }
            set
            {
                if (value)
                    CheckState = CheckState.Checked;
                else
                    CheckState = CheckState.Unchecked;
            }
        }

        public virtual bool IsLeaf
        {
            get { return false; }
        }

        #endregion

        public Node()
            : this(string.Empty)
        {
        }

        public Node(string text)
        {
            _text = text;
            _nodes = new NodeCollection(this);
        }

        public override string ToString()
        {
            return Text;
        }

        private TreeModel FindModel()
        {
            Node node = this;
            while (node != null)
            {
                if (node.Model != null)
                    return node.Model;
                node = node.Parent;
            }
            return null;
        }

        protected void NotifyModel()
        {
            TreeModel model = FindModel();
            if (model != null && Parent != null)
            {
                TreePath path = model.GetPath(Parent);
                if (path != null)
                {
                    var args = new TreeModelEventArgs(path, new[] {Index}, new object[] {this});
                    model.OnNodesChanged(args);
                }
            }
        }
    }
}