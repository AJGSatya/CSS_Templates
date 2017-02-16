using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CSSTemplates.TemplateSyncTool.Settings;
using log4net;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;

namespace CSSTemplates.TemplateSyncTool.Synchronise
{
    class ListSynchroniser : ItemSynchroniserBase
    {
        public override void HandleListItem(SharePointClientContext spContext, ILog logger, Arguments args, string destinationPath,
            ListItem item, ref int updatedCount, ref int noUpdateRequiredCount)
        {
            // nop
        }
    }
}
