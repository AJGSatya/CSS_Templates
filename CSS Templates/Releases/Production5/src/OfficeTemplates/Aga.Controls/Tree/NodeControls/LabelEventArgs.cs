using System;

namespace Aga.Controls.Tree.NodeControls
{
    public class LabelEventArgs : EventArgs
    {
        private readonly string _newLabel;
        private readonly string _oldLabel;
        private readonly object _subject;

        public LabelEventArgs(object subject, string oldLabel, string newLabel)
        {
            _subject = subject;
            _oldLabel = oldLabel;
            _newLabel = newLabel;
        }

        public object Subject
        {
            get { return _subject; }
        }

        public string OldLabel
        {
            get { return _oldLabel; }
        }

        public string NewLabel
        {
            get { return _newLabel; }
        }
    }
}