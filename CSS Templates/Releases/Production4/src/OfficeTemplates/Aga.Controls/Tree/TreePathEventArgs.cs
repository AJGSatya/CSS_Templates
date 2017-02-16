using System;

namespace Aga.Controls.Tree
{
    public class TreePathEventArgs : EventArgs
    {
        private readonly TreePath _path;

        public TreePathEventArgs()
        {
            _path = new TreePath();
        }

        public TreePathEventArgs(TreePath path)
        {
            if (path == null)
                throw new ArgumentNullException();

            _path = path;
        }

        public TreePath Path
        {
            get { return _path; }
        }
    }
}