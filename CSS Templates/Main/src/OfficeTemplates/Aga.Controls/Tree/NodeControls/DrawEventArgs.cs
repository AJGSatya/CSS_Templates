using System;
using System.Drawing;

namespace Aga.Controls.Tree.NodeControls
{
    public class DrawEventArgs : NodeEventArgs
    {
        private readonly DrawContext _context;
        private readonly EditableControl _control;

        public DrawEventArgs(TreeNodeAdv node, EditableControl control, DrawContext context, string text)
            : base(node)
        {
            _control = control;
            _context = context;
            Text = text;
        }

        public DrawContext Context
        {
            get { return _context; }
        }

        [Obsolete("Use TextColor")]
        public Brush TextBrush { get; set; }

        public Brush BackgroundBrush { get; set; }

        public Font Font { get; set; }

        public Color TextColor { get; set; }

        public string Text { get; set; }


        public EditableControl Control
        {
            get { return _control; }
        }
    }
}