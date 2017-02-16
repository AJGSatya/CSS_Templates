using System;
using System.IO;
using CommandLineParser.Arguments;

namespace CSSTemplates.TemplateSyncTool.Settings
{
    // https://github.com/j-maly/CommandLineParser
    class Arguments
    {
        [SwitchArgument('b', "breakonerror", false, Description = "Stops synchronisation on first error", Optional = true)]
        public bool BreakOnError { get; set; }
        [SwitchArgument('c', "clean", false, Description = "Cleans the target folder before downloading the fragments", Optional = true)]
        public bool Clean { get; set; }
        [SwitchArgument('h', "help", false, Description = "Displays this help text", Optional = true)]
        public bool Help { get; set; }
        [ValueArgument(typeof(string), 's', "source", Description = "Specifies the SharePoint source URL (e.g. https://ajgau.sharepoint.com/sites/Applications/Templates/)", Optional = true)]
        public string Source { get; set; }
        [SwitchArgument('o', "overwrite", false, Description = "Overwrites all local fragments with the SharePoint version", Optional = true)]
        public bool Overwrite { get; set; }
        [DirectoryArgument('p', "path", Description = "Specifies the target path where the fragments should be downloaded to (e.g. C:\\ProgramData\\AJG\\CSS Templates)", DirectoryMustExist = true, Optional = true)]
        public DirectoryInfo Directory { get; set; }
    }
}
