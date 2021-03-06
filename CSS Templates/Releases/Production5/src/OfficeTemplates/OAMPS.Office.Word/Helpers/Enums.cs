﻿namespace OAMPS.Office.Word.Helpers
{
    public static class Enums
    {
        public enum FormLoadType
        {
            OpenDocument,
            RibbonClick,
            RegenerateTemplate,
            NoWizard,
            ConvertWizard
        }

        public enum UsageTrackingType
        {
            NewDocument,
            UpdateData,
            RegenerateDocument,
            ConvertDocument
        }
    }
}