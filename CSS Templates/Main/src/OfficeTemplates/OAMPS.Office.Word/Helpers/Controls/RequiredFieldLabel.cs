using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace OAMPS.Office.Word.Helpers.Controls
{
    public class RequiredFieldLabel : Label
    {
        //[Localizable(true)]
        ////[System.Windows.Forms.SRDescription("ControlTextDescr")]
        ////[Bindable(true)]
        ////[System.Windows.Forms.SRCategory("CatAppearance")]
        //[DispId(-517)]
        [Browsable(true)]
        [Bindable(false)]
        [Localizable(true)]
        public Control Field
        {
            get { return Parent; }
            set { Parent = value; }
        }

        public override sealed string Text
        {
            get { return base.Text; }
            set { base.Text = @"*"; }
        }

        public override sealed Color ForeColor
        {
            get { return Color.Red; }
            set { base.ForeColor = Color.Red; }
        }
    }
}