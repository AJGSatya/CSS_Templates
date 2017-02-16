namespace CSSTemplates.TemplateSyncTool.Synchronise
{
    class UpdateCounter
    {
        public UpdateCounter()
        {
            DocumentCount = 0;
            DocumentUpdateCount = 0;
            ListItemCount = 0;
            ListItemUpdateCount = 0;
        }
        public int DocumentCount { get; set; }
        public int DocumentUpdateCount { get; set; }
        public int ListItemCount { get; set; }
        public int ListItemUpdateCount { get; set; }
    }
}