using System.IO;
using System.Windows.Forms;
using Aga.Controls.Properties;

namespace Aga.Controls
{
    public static class ResourceHelper
    {
        // VSpilt Cursor with Innerline (symbolisize hidden column)
        private static readonly Cursor _dVSplitCursor = GetCursor(Resources.DVSplit);

        private static readonly GifDecoder _loadingIcon = GetGifDecoder(Resources.loading_icon);

        public static Cursor DVSplitCursor
        {
            get { return _dVSplitCursor; }
        }

        public static GifDecoder LoadingIcon
        {
            get { return _loadingIcon; }
        }

        /// <summary>
        ///     Help function to convert byte[] from resource into Cursor Type
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static Cursor GetCursor(byte[] data)
        {
            using (var s = new MemoryStream(data))
                return new Cursor(s);
        }

        /// <summary>
        ///     Help function to convert byte[] from resource into GifDecoder Type
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static GifDecoder GetGifDecoder(byte[] data)
        {
            using (var ms = new MemoryStream(data))
                return new GifDecoder(ms, true);
        }
    }
}