﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace TalkdeskReportGenerator.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "15.8.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("AH")]
        public string PhoneColorKeyColumn {
            get {
                return ((string)(this["PhoneColorKeyColumn"]));
            }
            set {
                this["PhoneColorKeyColumn"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("4")]
        public int PhoneColorKeyRow {
            get {
                return ((int)(this["PhoneColorKeyRow"]));
            }
            set {
                this["PhoneColorKeyRow"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("F")]
        public string GroupByNameColumn {
            get {
                return ((string)(this["GroupByNameColumn"]));
            }
            set {
                this["GroupByNameColumn"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("12")]
        public int GroupByNameRow {
            get {
                return ((int)(this["GroupByNameRow"]));
            }
            set {
                this["GroupByNameRow"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("G")]
        public string AgentNameColumn {
            get {
                return ((string)(this["AgentNameColumn"]));
            }
            set {
                this["AgentNameColumn"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("H")]
        public string TwelveAmColumn {
            get {
                return ((string)(this["TwelveAmColumn"]));
            }
            set {
                this["TwelveAmColumn"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Central Standard Time")]
        public string TimeZoneId {
            get {
                return ((string)(this["TimeZoneId"]));
            }
            set {
                this["TimeZoneId"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("C:\\")]
        public string OutputDirectory {
            get {
                return ((string)(this["OutputDirectory"]));
            }
            set {
                this["OutputDirectory"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public global::System.Collections.Generic.List<System.String> TemporaryExcelPaths {
            get {
                return ((global::System.Collections.Generic.List<System.String>)(this["TemporaryExcelPaths"]));
            }
            set {
                this["TemporaryExcelPaths"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("C:\\")]
        public string InputDirectory {
            get {
                return ((string)(this["InputDirectory"]));
            }
            set {
                this["InputDirectory"] = value;
            }
        }
    }
}
