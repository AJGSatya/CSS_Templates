using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aga.Controls.Tree;

namespace OAMPS.Office.Word.Helpers.Controls
{
   public class AdvancedTreeNode : Node
    {
 /// <exception cref="ArgumentNullException">Argument is null.</exception>
	    public override string Text
		{
			get
			{
				return base.Text;
			}
			set
			{
				if (string.IsNullOrEmpty(value))
					throw new ArgumentNullException();

				base.Text = value;
			}
		}

       /// <summary>
       /// Whether the box is checked or not.
       /// </summary>
       public bool Checked { get; set; }

       public string Current { get; set; }
       public string Reccommended { get; set; }

       /// <summary>
        /// Initializes a new MyNode class with a given Text property.
        /// </summary>
        /// <param name="text">String to set the text property with.</param>
        public AdvancedTreeNode(string text)
			: base(text)
		{
		}
    }
}
