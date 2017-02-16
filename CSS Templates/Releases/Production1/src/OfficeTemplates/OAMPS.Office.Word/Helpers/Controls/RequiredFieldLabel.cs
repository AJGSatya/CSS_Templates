using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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
            get { return base.Parent; } 
            set { base.Parent = value; }   
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

        //public RequiredFieldLabel()
        //{
          
        //    //if(base.Parent != null)
        //    //    base.Location = new Point(base.Parent.Location.X + 5, base.Parent.Location.Y);
        //}

        //public override sealed Point Location
        //{
        //    get { return base.Location; }
        //    set { base.SetBounds(Field.Location.X + Field.Width, Field.Location.Y, base.Width, this.Height, BoundsSpecified.Location); }
        //}
    }
}
