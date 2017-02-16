using System;
using System.Configuration;
using System.IO;
using System.Linq.Expressions;
using CommandLineParser.Arguments;
using CommandLineParser.Exceptions;
using CSSTemplates.TemplateSyncTool.Properties;
using CSSTemplates.TemplateSyncTool.Settings;
using CSSTemplates.TemplateSyncTool.Synchronise;
using Microsoft.SharePoint.Client;

namespace CSSTemplates.TemplateSyncTool
{
    class TemplateSyncTool
    {
        private static log4net.ILog _logger;

        static void Main(string[] args)
        {
            log4net.Config.BasicConfigurator.Configure();
            _logger = log4net.LogManager.GetLogger(typeof(TemplateSyncTool));

            var arguments = ParseArguments(args);
            if (arguments != null)
            {
                var sync = new Synchroniser();
                sync.Synchronize(_logger, arguments);
            }
        }

        private static Arguments ParseArguments(string[] args)
        {
            var parser = new CommandLineParser.CommandLineParser();
            var arguments = new Arguments();
            parser.ExtractArgumentAttributes(arguments);

            try
            {
                parser.ParseCommandLine(args);

                if (arguments.Directory == null)
                {
                    arguments.Directory = new DirectoryInfo(ConfigurationManager.AppSettings["TemplateDownloadDirectory"]);
                }

                if (string.IsNullOrEmpty(arguments.Source))
                {
                    arguments.Source = ConfigurationManager.AppSettings["SharePointContextUrl"];
                }

                if (!CheckURLValid(arguments.Source))
                {
                    Console.Error.WriteLine("ERROR: {0} is not a valid URL", arguments.Source);
                    parser.ShowUsage();
                    return null;
                }

                if (arguments.Help)
                {
                    parser.ShowUsage();
                    return null;
                }
            }
            catch (Exception)
            {
                parser.ShowUsage();
                return null;
            }

            return arguments;
        }

        private static bool CheckURLValid(string source)
        {
            Uri uriResult;
            return Uri.TryCreate(source, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
        }

    }
}
