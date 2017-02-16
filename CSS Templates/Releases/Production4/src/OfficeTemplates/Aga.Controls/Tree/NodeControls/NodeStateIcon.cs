using System.Drawing;
using Aga.Controls.Properties;

namespace Aga.Controls.Tree.NodeControls
{
    public class NodeStateIcon : NodeIcon
    {
        private readonly Image _closed;
        private readonly Image _leaf;
        private readonly Image _opened;

        public NodeStateIcon()
        {
            _leaf = MakeTransparent(Resources.Leaf);
            _opened = MakeTransparent(Resources.Folder);
            _closed = MakeTransparent(Resources.FolderClosed);
        }

        private static Image MakeTransparent(Bitmap bitmap)
        {
            bitmap.MakeTransparent(bitmap.GetPixel(0, 0));
            return bitmap;
        }

        protected override Image GetIcon(TreeNodeAdv node)
        {
            Image icon = base.GetIcon(node);
            if (icon != null)
                return icon;
            else if (node.IsLeaf)
                return _leaf;
            else if (node.CanExpand && node.IsExpanded)
                return _opened;
            else
                return _closed;
        }
    }
}