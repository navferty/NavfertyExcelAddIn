﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Navferty.Common.Localization {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class UIStrings {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal UIStrings() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Navferty.Common.Localization.UIStrings", typeof(UIStrings).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Failed to send feedback email!.
        /// </summary>
        internal static string Feedback_ErrorTitle {
            get {
                return ResourceManager.GetString("Feedback_ErrorTitle", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Bug tracker and feature request.
        /// </summary>
        internal static string Feedback_GotoGithub {
            get {
                return ResourceManager.GetString("Feedback_GotoGithub", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Include screenshots.
        /// </summary>
        internal static string Feedback_IncludeScreenshots {
            get {
                return ResourceManager.GetString("Feedback_IncludeScreenshots", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Briefly describe the error that occurs (no more than {0} characters):.
        /// </summary>
        internal static string Feedback_Message {
            get {
                return ResourceManager.GetString("Feedback_Message", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Send.
        /// </summary>
        internal static string Feedback_Send {
            get {
                return ResourceManager.GetString("Feedback_Send", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to The error log will be included in the report..
        /// </summary>
        internal static string Feedback_Summary {
            get {
                return ResourceManager.GetString("Feedback_Summary", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Feedback.
        /// </summary>
        internal static string Feedback_Title {
            get {
                return ResourceManager.GetString("Feedback_Title", resourceCulture);
            }
        }
    }
}
