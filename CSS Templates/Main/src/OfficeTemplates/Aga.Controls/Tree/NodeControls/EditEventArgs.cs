﻿using System.Windows.Forms;

namespace Aga.Controls.Tree.NodeControls
{
    public class EditEventArgs : NodeEventArgs
    {
        private readonly Control _control;

        public EditEventArgs(TreeNodeAdv node, Control control)
            : base(node)
        {
            _control = control;
        }

        public Control Control
        {
            get { return _control; }
        }
    }
}