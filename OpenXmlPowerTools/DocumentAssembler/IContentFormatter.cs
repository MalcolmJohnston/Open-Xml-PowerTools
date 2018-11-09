using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Interface for Document Assembler content formatters
    /// Allows for extending the way that content is presented in OpenXML
    /// based on source format.
    /// </summary>
    public interface IContentFormatter
    {
        /// <summary>
        /// Adds text content from an XML source file to the current paragraph
        /// in an OpenXML document.
        /// Formatting for the text content should be processed here and applied to the 
        /// output OpenXML
        /// </summary>
        /// <param name="para">The current paragraph.</param>
        /// <param name="run">The current run.</param>
        /// <param name="newContent">The new content.</param>
        /// <returns>A collection of XElements representing the formatted OpenXML content.</returns>
        IEnumerable<XElement> FormatContent(XElement para, XElement run, string newContent);
    }
}
