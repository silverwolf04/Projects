﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace PdfCreator.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.7.0.0")]
    internal sealed partial class FacStaffPdf : global::System.Configuration.ApplicationSettingsBase {
        
        private static FacStaffPdf defaultInstance = ((FacStaffPdf)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new FacStaffPdf())));
        
        public static FacStaffPdf Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=csmis3; Initial Catalog=peopletest; Integrated Security=true")]
        public string ConnectionString {
            get {
                return ((string)(this["ConnectionString"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("select top 100 workphone, LName + \',\' + Fname as EmpName, department, campusBuild" +
            "ing, campusOfficeNo, email, faxphone, title from fac_staff_dir_view1 where relea" +
            "seflag is null order by LName, fname")]
        public string QueryString {
            get {
                return ((string)(this["QueryString"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("MSSQL")]
        public string DataProvider {
            get {
                return ((string)(this["DataProvider"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("EmpName")]
        public string Name {
            get {
                return ((string)(this["Name"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Workphone")]
        public string PhoneNumber {
            get {
                return ((string)(this["PhoneNumber"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Department")]
        public string Department {
            get {
                return ((string)(this["Department"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("FaxPhone")]
        public string FaxNumber {
            get {
                return ((string)(this["FaxNumber"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Email")]
        public string EmailAddress {
            get {
                return ((string)(this["EmailAddress"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Title")]
        public string Title {
            get {
                return ((string)(this["Title"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("CampusBuilding")]
        public string Building {
            get {
                return ((string)(this["Building"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("CampusOfficeNo")]
        public string Office {
            get {
                return ((string)(this["Office"]));
            }
        }
    }
}
