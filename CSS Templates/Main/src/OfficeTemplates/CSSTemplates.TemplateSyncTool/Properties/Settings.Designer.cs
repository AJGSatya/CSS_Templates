﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CSSTemplates.TemplateSyncTool.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "14.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("ajgau.onmicrosoft.com")]
        public string tenant {
            get {
                return ((string)(this["tenant"]));
            }
            set {
                this["tenant"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("83141754-35c8-4bf9-9a1d-c125529856f0")]
        public string clientid {
            get {
                return ((string)(this["clientid"]));
            }
            set {
                this["clientid"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://ajgau.sharepoint.com/sites/intranet")]
        public string redirecturi {
            get {
                return ((string)(this["redirecturi"]));
            }
            set {
                this["redirecturi"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://ajgau.sharepoint.com")]
        public string TokenResourceUrl {
            get {
                return ((string)(this["TokenResourceUrl"]));
            }
            set {
                this["TokenResourceUrl"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://login.microsoftonline.com/4f955c1f-62e5-44ef-8da3-051caaa6b103/federation" +
            "metadata/2007-06/federationmetadata.xml")]
        public string TenantMetadataUrl {
            get {
                return ((string)(this["TenantMetadataUrl"]));
            }
            set {
                this["TenantMetadataUrl"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("https://ajgau.sharepoint.com/sites/Applications/Templates/")]
        public string SharePointContextUrl {
            get {
                return ((string)(this["SharePointContextUrl"]));
            }
            set {
                this["SharePointContextUrl"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("OAMPSWordTemplates")]
        public string TemplateLibraryName {
            get {
                return ((string)(this["TemplateLibraryName"]));
            }
            set {
                this["TemplateLibraryName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("C:\\ProgramData\\AJG\\CSS Templates")]
        public string TemplateDownloadDirectory {
            get {
                return ((string)(this["TemplateDownloadDirectory"]));
            }
            set {
                this["TemplateDownloadDirectory"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("OAMPS Letters")]
        public string LetterTemplateLibraryName {
            get {
                return ((string)(this["LetterTemplateLibraryName"]));
            }
            set {
                this["LetterTemplateLibraryName"] = value;
            }
        }
    }
}
