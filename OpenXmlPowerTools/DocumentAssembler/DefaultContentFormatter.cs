using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Default implementation of IContentFormatter
    /// This is taken verbatim from the previous DocumentAssembler source.
    /// </summary>
    /// <seealso cref="OpenXmlPowerTools.DocumentAssembler.IContentFormatter" />
    public class DefaultContentFormatter : IContentFormatter
    {
        /// <summary>
        /// Formats the content.
        /// </summary>
        /// <param name="para">The current paragraph.</param>
        /// <param name="newContent">The new content.</param>
        /// <returns></returns>
        public IEnumerable<XElement> FormatContent(XElement para, XElement run, string newContent)
        {
            if (para != null)
            {
                XElement p = new XElement(W.p, para.Elements(W.pPr));
                foreach (string line in newContent.Split('\n'))
                {
                    p.Add(new XElement(W.r,
                            para.Elements(W.r).Elements(W.rPr).FirstOrDefault(),
                        (p.Elements().Count() > 1) ? new XElement(W.br) : null,
                        new XElement(W.t, line)));
                }

                return new XElement[] { p };
            }
            else
            {
                List<XElement> list = new List<XElement>();
                foreach (string line in newContent.Split('\n'))
                {
                    list.Add(new XElement(W.r,
                        run.Elements().Where(e => e.Name != W.t),
                        (list.Count > 0) ? new XElement(W.br) : null,
                        new XElement(W.t, line)));
                }

                return list;
            }
        }
    }
}
