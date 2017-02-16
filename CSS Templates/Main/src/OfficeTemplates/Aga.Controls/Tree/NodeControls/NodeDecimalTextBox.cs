using System.ComponentModel;
using System.Windows.Forms;

namespace Aga.Controls.Tree.NodeControls
{
    public class NodeDecimalTextBox : NodeTextBox
    {
        private bool _allowDecimalSeparator = true;

        private bool _allowNegativeSign = true;

        protected NodeDecimalTextBox()
        {
        }

        [DefaultValue(true)]
        public bool AllowDecimalSeparator
        {
            get { return _allowDecimalSeparator; }
            set { _allowDecimalSeparator = value; }
        }

        [DefaultValue(true)]
        public bool AllowNegativeSign
        {
            get { return _allowNegativeSign; }
            set { _allowNegativeSign = value; }
        }

        protected override TextBox CreateTextBox()
        {
            var textBox = new NumericTextBox();
            textBox.AllowDecimalSeparator = AllowDecimalSeparator;
            textBox.AllowNegativeSign = AllowNegativeSign;
            return textBox;
        }

        protected override void DoApplyChanges(TreeNodeAdv node, Control editor)
        {
            SetValue(node, (editor as NumericTextBox).DecimalValue);
        }
    }
}