using System;
using Aga.Controls.Tree;

namespace OAMPS.Office.Word.Helpers.Controls
{
    public class AdvancedTreeNode : Node
    {
        /// <summary>
        ///     Initializes a new MyNode class with a given Text property.
        /// </summary>
        /// <param name="text">String to set the text property with.</param>
        public AdvancedTreeNode(string text)
            : base(text)
        {
        }

        /// <exception cref="ArgumentNullException">Argument is null.</exception>
        public override string Text
        {
            get { return base.Text; }
            set
            {
                if (string.IsNullOrEmpty(value))
                    throw new ArgumentNullException();

                base.Text = value;
            }
        }

        /// <summary>
        ///     Whether the box is checked or not.
        /// </summary>
        public bool Checked { get; set; }
        public string Id { get; set; }
        public string Current { get; set; }
        public string Reccommended { get; set; }

        public string CurrentId { get; set; }
        public string ReccommendedId { get; set; }
        public string OrderPolicy { get; set; }

        public string PolicyNumber { get; set; }

        public string Insurer { get; set; }
        public string InsurerId { get; set; }
    }
}