using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Aga.Controls.Properties;

namespace Aga.Controls.Tree.NodeControls
{
    internal class NodePlusMinus : NodeControl
    {
        public const int ImageSize = 9;
        public const int Width = 16;
        private readonly Bitmap _minus;
        private readonly Bitmap _plus;
        private VisualStyleRenderer _closedRenderer;

        private VisualStyleRenderer _openedRenderer;

        public NodePlusMinus()
        {
            _plus = Resources.plus;
            _minus = Resources.minus;
        }

        private VisualStyleRenderer OpenedRenderer
        {
            get
            {
                if (_openedRenderer == null)
                    _openedRenderer = new VisualStyleRenderer(VisualStyleElement.TreeView.Glyph.Opened);
                return _openedRenderer;
            }
        }

        private VisualStyleRenderer ClosedRenderer
        {
            get
            {
                if (_closedRenderer == null)
                    _closedRenderer = new VisualStyleRenderer(VisualStyleElement.TreeView.Glyph.Closed);
                return _closedRenderer;
            }
        }

        public override Size MeasureSize(TreeNodeAdv node, DrawContext context)
        {
            return new Size(Width, Width);
        }

        public override void Draw(TreeNodeAdv node, DrawContext context)
        {
            if (node.CanExpand)
            {
                Rectangle r = context.Bounds;
                var dy = (int) Math.Round((float) (r.Height - ImageSize)/2);
                if (Application.RenderWithVisualStyles)
                {
                    VisualStyleRenderer renderer;
                    if (node.IsExpanded)
                        renderer = OpenedRenderer;
                    else
                        renderer = ClosedRenderer;
                    renderer.DrawBackground(context.Graphics, new Rectangle(r.X, r.Y + dy, ImageSize, ImageSize));
                }
                else
                {
                    Image img;
                    if (node.IsExpanded)
                        img = _minus;
                    else
                        img = _plus;
                    context.Graphics.DrawImageUnscaled(img, new Point(r.X, r.Y + dy));
                }
            }
        }

        public override void MouseDown(TreeNodeAdvMouseEventArgs args)
        {
            if (args.Button == MouseButtons.Left)
            {
                args.Handled = true;
                if (args.Node.CanExpand)
                    args.Node.IsExpanded = !args.Node.IsExpanded;
            }
        }

        public override void MouseDoubleClick(TreeNodeAdvMouseEventArgs args)
        {
            args.Handled = true; // Supress expand/collapse when double click on plus/minus
        }
    }
}