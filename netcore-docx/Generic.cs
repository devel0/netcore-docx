using System;
using System.Linq;
using static System.Math;
using System.Collections.Generic;
using System.Globalization;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using SearchAThing;
using SearchAThing.DocX;
using static SearchAThing.DocX.Constants;
using System.Runtime.CompilerServices;

namespace SearchAThing.DocX
{

    public static partial class DocXExt
    {

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static MainDocumentPart GetMainDocumentPart(this OpenXmlElement element)
        {
            while (element.Parent is not null) element = element.Parent;

            if (element is Document document)
                return document.MainDocumentPart!;

            else if (element is Header header)
                return ((WordprocessingDocument)header.HeaderPart!.OpenXmlPackage).MainDocumentPart!;

            else if (element is Footer footer)
                return ((WordprocessingDocument)footer.FooterPart!.OpenXmlPackage).MainDocumentPart!;

            else
                throw new Exception($"openxmlelement with topmost parent [{element.GetType()}]");
        }

        /// <summary>
        /// retrieve WordprocessingDocument from openxmlelement walking to topmost document and from there to
        /// openxmlpackage (WordprocessingDocument) through MainDocumentPart
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static WordprocessingDocument GetWordprocessingDocument(this OpenXmlElement element)
        {
            while (element.Parent is not null) element = element.Parent;

            if (element is Document document)
                return (WordprocessingDocument)document.MainDocumentPart!.OpenXmlPackage;

            else if (element is Header header)
                return (WordprocessingDocument)header.HeaderPart!.OpenXmlPackage;

            else if (element is Footer footer)
                return (WordprocessingDocument)footer.FooterPart!.OpenXmlPackage;

            else if (element is Paragraph paragraph)
                throw new Exception($"paragraph will null parent");

            else
                throw new Exception($"openxmlelement with topmost parent [{element.GetType()}]");
        }

        /// <summary>
        /// retrieve body from OpenXmlElement walking through parent
        /// </summary>        
        public static Body GetBody(this OpenXmlElement element)
        {
            if (element is Body body) return body;

            if (element.Parent is null) throw new Exception($"could'n find body");

            return element.Parent.GetBody();
        }

        /// <summary>
        /// retrieve typed element type child from the given owner element.
        /// if createIfNotExists was true and child not exists a new child will inserted.
        /// by default new child will append or if insertAtIdx specified chilld will inserted at the given index.
        /// </summary>
        /// <param name="element">element owner</param>
        /// <param name="createIfNotExists">(optional) if true and child not found a new one will created</param>
        /// <param name="insertAtIdx">(optional) if specified new child will inserted at given index</param>
        /// <param name="onNew">(optional) custom action to apply to child if it is created because missing</param>
        /// <returns>existing or new child</returns>        
        public static T? GetOrCreate<T>(this OpenXmlElement element,
            bool createIfNotExists,
            int? insertAtIdx = null,
            Action<T>? onNew = null)
            where T : OpenXmlElement, new()
        {
            var res = element.Elements<T>().FirstOrDefault();

            if (res is null && createIfNotExists)
            {
                res = new T();
                if (insertAtIdx is not null)
                    element.InsertAt(res, insertAtIdx.Value);
                else
                    element.Append(res);
                onNew?.Invoke(res);
            }

            return res;
        }

        internal static WrapperManager GetWrapperRef<T>(this T element) where T : OpenXmlElement =>
            WrapperManager.GetWrapperRef(element.GetWordprocessingDocument());

        public static T EnableAutoNumbering<T>(this T element,
            NumberFormatValues type = NumberFormatValues.Bullet,
            int level = 0,
            bool structured = false) where T : OpenXmlElement
        {
            var wrapper = element.GetWrapperRef();

            wrapper.numberingEnabled = type;
            wrapper.numberingLevel = level;
            wrapper.numberingStructured = structured;

            return element;
        }

        public static T IncAutoNumbering<T>(this T element,
            NumberFormatValues type = NumberFormatValues.Bullet,
            int level = 0,
            bool structured = false) where T : OpenXmlElement
        {
            var wrapper = element.GetWrapperRef();

            if (wrapper.numberingLevel < 10)
                ++wrapper.numberingLevel;

            return element;
        }

        public static T DecAutoNumbering<T>(this T element,
           NumberFormatValues type = NumberFormatValues.Bullet,
           int level = 0,
           bool structured = false) where T : OpenXmlElement
        {
            var wrapper = element.GetWrapperRef();

            if (wrapper.numberingLevel > 0)
                --wrapper.numberingLevel;

            return element;
        }

        public static T DisableAutoNumbering<T>(this T element) where T : OpenXmlElement
        {
            var wrapper = element.GetWrapperRef();

            wrapper.numberingEnabled = null;

            return element;
        }

        /// <summary>
        /// retrieve last element that isn't main section property
        /// </summary>
        /// <param name="doc">wordprocessing doc</param>
        /// <returns>element</returns>
        public static OpenXmlElement? GetLastElement(this WordprocessingDocument doc)
        {
            var body = doc.GetBody();

            var element = body.LastChild;

            if (!(element is SectionProperties)) throw new Exception($"missing main section properties");

            element = element.PreviousSibling();

            if (element is null) return null;

            return element;
        }

        /// <summary>
        /// retrieve next element of this by specified given type T and optional condition
        /// </summary>
        /// <param name="element">element from where start search next</param>
        /// <param name="condition">condition to eval next element as result candidate</param>        
        /// <returns>next element of this with type T and optional condition</returns>
        public static T? NextSibling<T>(this OpenXmlElement element,
            Func<OpenXmlElement, bool>? condition = null) where T : OpenXmlElement
        {
            var nextElement = element.NextSibling<T>();

            if (nextElement is null) return null;

            if (condition is not null)
            {
                if (condition(nextElement)) return nextElement;

                return nextElement.NextSibling<T>(condition);
            }

            return nextElement;
        }

        /// <summary>
        /// retrieve next element of this by optional condition
        /// </summary>
        /// <param name="element">element from where start search next</param>
        /// <param name="condition">condition to eval next element as result candidate</param>        
        /// <returns>next element of this with optional condition</returns>
        public static OpenXmlElement? NextSibling(this OpenXmlElement element,
            Func<OpenXmlElement, bool>? condition = null)
        {
            var nextElement = element.NextSibling();

            if (nextElement is null) return null;

            if (condition is not null)
            {
                if (condition(nextElement)) return nextElement;

                return nextElement.NextSibling(condition);
            }

            return nextElement;
        }

        /// <summary>
        /// add given customParagraph ( it will automatically detached from its parent if not already null )
        /// </summary>
        /// <param name="elementBefore">elemento to which append custom paragraph</param>
        /// <param name="customParagraph">custom paragraph</param>
        /// <returns>custom paragraph</returns>
        public static Paragraph AddParagraph(this OpenXmlElement elementBefore,
            Paragraph customParagraph)
        {
            if (customParagraph.Parent is not null)
            {
                customParagraph.Remove();
            }

            if (elementBefore.Parent is null) throw new ArgumentException($"given element before has null Parent");

            elementBefore.Parent.Append(customParagraph);

            return customParagraph;
        }

        /// <summary>
        /// add a new paragraph after given paragraphBefore or to end of document
        /// </summary>
        /// <param name="doc">word processing document</param>
        /// <param name="txt">paragraph initial text</param>
        /// <param name="paragraphBefore">(optional) paragraph before the new one</param>
        /// <param name="elementBefore">(optional) element before the new one</param>
        /// <param name="style">(optional) style to apply to new paragraph or if exists a previous one it will inherithed</param>
        /// <param name="action">(optional) execute action on run created with this paragrah if any</param>
        /// <returns>new paragraph</returns>
        public static Paragraph AddParagraph(this OpenXmlElement elementBefore,
            string? txt = null,
            Style? style = null,
            Paragraph? paragraphBefore = null,
            Action<Run>? action = null,
            WordprocessingDocument? doc = null)
        {
            if (doc is null) doc = elementBefore.GetWordprocessingDocument();

            OpenXmlElement? parent = null;

            if (parent is null)
            {
                if (elementBefore is Header || elementBefore is Footer || elementBefore is TableCell)
                    parent = elementBefore;
            }

            return doc.AddParagraph(
                txt: txt,
                paragraphBefore: paragraphBefore,
                elementBefore: elementBefore,
                parent: parent,
                style: style,
                action: action);
        }


        /// <summary>
        /// add a break ( page, column or textwrapping )
        /// </summary>
        /// <param name="element">element after which apply the break</param>
        /// <param name="type">type of break</param>
        /// <returns>next paragraph with break applied</returns>
        public static Paragraph AddBreak<T>(this T element, BreakValues type = BreakValues.Page) where T : OpenXmlElement
        {
            var doc = element.GetWordprocessingDocument();

            if (element is Paragraph paragraph)
                return doc.AddParagraph("",
                    paragraphBefore: paragraph,
                    elementBefore: element,
                    action: run => run.Append(new Break { Type = type }));

            else
                return doc.AddParagraph("",
                    elementBefore: element,
                    action: run => run.Append(new Break { Type = type }));
        }


        internal static T SetShading<T, P>(this T element,
            System.Drawing.Color? color = null,
            ShadingPatternValues pattern = ShadingPatternValues.Clear)
            where P : OpenXmlElement, new()
            where T : OpenXmlElement
        {
            var runProperties = element.GetOrCreate<P>(createIfNotExists: true, insertAtIdx: 0)!;
            var shading = runProperties.GetOrCreate<Shading>(createIfNotExists: false)!;

            if (shading is null)
            {
                if (color is not null)
                {
                    shading = createShading(runProperties, color, pattern);
                }
            }
            else
            {
                if (color is null)
                    shading.Remove();

                else
                    shading = createShading(runProperties, color, pattern);
            }

            return element;
        }

        static Shading createShading(OpenXmlElement runOrParagraphProperties,
            System.Drawing.Color? color,
            ShadingPatternValues pattern)
        {
            var shading = runOrParagraphProperties.GetOrCreate<Shading>(createIfNotExists: true)!;
            shading.Val = pattern;
            shading.Fill = color.ToWPColorString();

            return shading;
        }


    }

}