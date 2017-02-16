using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Aga.Controls.Tree.NodeControls
{
    public class NodeNumericUpDown : BaseTextControl
    {
        #region Properties

        private int _editorWidth = 100;
        private decimal _increment = 1;
        private decimal _maximum = 100;

        [DefaultValue(100)]
        public int EditorWidth
        {
            get { return _editorWidth; }
            set { _editorWidth = value; }
        }

        [Category("Data"), DefaultValue(0)]
        public int DecimalPlaces { get; set; }

        [Category("Data"), DefaultValue(1)]
        public decimal Increment
        {
            get { return _increment; }
            set { _increment = value; }
        }

        [Category("Data"), DefaultValue(0)]
        public decimal Minimum { get; set; }

        [Category("Data"), DefaultValue(100)]
        public decimal Maximum
        {
            get { return _maximum; }
            set { _maximum = value; }
        }

        #endregion

        protected override Size CalculateEditorSize(EditorContext context)
        {
            if (Parent.UseColumns)
                return context.Bounds.Size;
            else
                return new Size(EditorWidth, context.Bounds.Height);
        }

        protected override Control CreateEditor(TreeNodeAdv node)
        {
            var num = new NumericUpDown();
            num.Increment = Increment;
            num.DecimalPlaces = DecimalPlaces;
            num.Minimum = Minimum;
            num.Maximum = Maximum;
            num.Value = (decimal) GetValue(node);
            SetEditControlProperties(num, node);
            return num;
        }

        protected override void DisposeEditor(Control editor)
        {
        }

        protected override void DoApplyChanges(TreeNodeAdv node, Control editor)
        {
            SetValue(node, (editor as NumericUpDown).Value);
        }
    }
}