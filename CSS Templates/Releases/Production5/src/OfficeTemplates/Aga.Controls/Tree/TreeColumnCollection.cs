using System;
using System.Collections.ObjectModel;
using System.Windows.Forms;

namespace Aga.Controls.Tree
{
    internal class TreeColumnCollection : Collection<TreeColumn>
    {
        private readonly TreeViewAdv _treeView;

        public TreeColumnCollection(TreeViewAdv treeView)
        {
            _treeView = treeView;
        }

        protected override void InsertItem(int index, TreeColumn item)
        {
            base.InsertItem(index, item);
            BindEvents(item);
            _treeView.UpdateColumns();
        }

        protected override void RemoveItem(int index)
        {
            UnbindEvents(this[index]);
            base.RemoveItem(index);
            _treeView.UpdateColumns();
        }

        protected override void SetItem(int index, TreeColumn item)
        {
            UnbindEvents(this[index]);
            base.SetItem(index, item);
            item.Owner = this;
            BindEvents(item);
            _treeView.UpdateColumns();
        }

        protected override void ClearItems()
        {
            foreach (TreeColumn c in Items)
                UnbindEvents(c);
            Items.Clear();
            _treeView.UpdateColumns();
        }

        private void BindEvents(TreeColumn item)
        {
            item.Owner = this;
            item.HeaderChanged += HeaderChanged;
            item.IsVisibleChanged += IsVisibleChanged;
            item.WidthChanged += WidthChanged;
            item.SortOrderChanged += SortOrderChanged;
        }

        private void UnbindEvents(TreeColumn item)
        {
            item.Owner = null;
            item.HeaderChanged -= HeaderChanged;
            item.IsVisibleChanged -= IsVisibleChanged;
            item.WidthChanged -= WidthChanged;
            item.SortOrderChanged -= SortOrderChanged;
        }

        private void SortOrderChanged(object sender, EventArgs e)
        {
            var changed = sender as TreeColumn;
            //Only one column at a time can have a sort property set
            if (changed.SortOrder != SortOrder.None)
            {
                foreach (TreeColumn col in this)
                {
                    if (col != changed)
                        col.SortOrder = SortOrder.None;
                }
            }
            _treeView.UpdateHeaders();
        }

        private void WidthChanged(object sender, EventArgs e)
        {
            _treeView.ChangeColumnWidth(sender as TreeColumn);
        }

        private void IsVisibleChanged(object sender, EventArgs e)
        {
            _treeView.FullUpdate();
        }

        private void HeaderChanged(object sender, EventArgs e)
        {
            _treeView.UpdateView();
        }
    }
}