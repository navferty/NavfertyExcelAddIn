﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.42000
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace NavfertyExcelAddIn.Localization {
    using System;
    
    
    /// <summary>
    ///   Класс ресурса со строгой типизацией для поиска локализованных строк и т.д.
    /// </summary>
    // Этот класс создан автоматически классом StronglyTypedResourceBuilder
    // с помощью такого средства, как ResGen или Visual Studio.
    // Чтобы добавить или удалить член, измените файл .ResX и снова запустите ResGen
    // с параметром /str или перестройте свой проект VS.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class RibbonSupertips {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal RibbonSupertips() {
        }
        
        /// <summary>
        ///   Возвращает кэшированный экземпляр ResourceManager, использованный этим классом.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("NavfertyExcelAddIn.Localization.RibbonSupertips", typeof(RibbonSupertips).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Перезаписывает свойство CurrentUICulture текущего потока для всех
        ///   обращений к ресурсу с помощью этого класса ресурса со строгой типизацией.
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
        ///   Ищет локализованную строку, похожую на Saves the selected range of values to the clipboard using Markdown markup..
        /// </summary>
        internal static string CopyAsMarkdown {
            get {
                return ResourceManager.GetString("CopyAsMarkdown", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на In the development....
        /// </summary>
        internal static string CutNames {
            get {
                return ResourceManager.GetString("CutNames", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Finds all cells in the specified range that have any errors (#)..
        /// </summary>
        internal static string FindErrors {
            get {
                return ResourceManager.GetString("FindErrors", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Paints groups of duplicates with different colors..
        /// </summary>
        internal static string HighlightDuplicates {
            get {
                return ResourceManager.GetString("HighlightDuplicates", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Converts the cell data type to a numeric one..
        /// </summary>
        internal static string ParseNumerics {
            get {
                return ResourceManager.GetString("ParseNumerics", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Changes similar letters of the Russian and English alphabets, for example: &apos;У&apos; - &gt; &apos;Y&apos;..
        /// </summary>
        internal static string ReplaceChars {
            get {
                return ResourceManager.GetString("ReplaceChars", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Rewrites the numeric value of the cell with words. The default is the Russian language..
        /// </summary>
        internal static string StringifyNumericsButton {
            get {
                return ResourceManager.GetString("StringifyNumericsButton", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Rewrites the numeric value of a cell in English with words..
        /// </summary>
        internal static string StringifyNumericsEn {
            get {
                return ResourceManager.GetString("StringifyNumericsEn", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Rewrites the numeric value of a cell in Russian with words..
        /// </summary>
        internal static string StringifyNumericsRu {
            get {
                return ResourceManager.GetString("StringifyNumericsRu", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Changes the case of all words in the cell. Replacement order: lowercase, uppercase, first capital, lowercase....
        /// </summary>
        internal static string ToggleCase {
            get {
                return ResourceManager.GetString("ToggleCase", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Changes the Russian alphabet to English using transliteration, for example: &apos;У&apos; - &gt; &apos;U&apos;..
        /// </summary>
        internal static string Transliterate {
            get {
                return ResourceManager.GetString("Transliterate", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Full using transliteration (by default) or partial replacement of the Russian alphabet (by analogy) with English. Example of transliteration: &apos;У&apos; - &gt; &apos;U&apos;, example of substitution by analogy:&apos; У &apos; - &gt; &apos;Y&apos;..
        /// </summary>
        internal static string TransliterateButton {
            get {
                return ResourceManager.GetString("TransliterateButton", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Reduces the groups of spaces between the entered parts of the cell value to one..
        /// </summary>
        internal static string TrimSpaces {
            get {
                return ResourceManager.GetString("TrimSpaces", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Parts of the cell duplicate their content..
        /// </summary>
        internal static string UnmergeCells {
            get {
                return ResourceManager.GetString("UnmergeCells", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Ищет локализованную строку, похожую на Checking that values meet certain standards. Use the pop-up list to select a standard..
        /// </summary>
        internal static string ValidateValuesButton {
            get {
                return ResourceManager.GetString("ValidateValuesButton", resourceCulture);
            }
        }
    }
}