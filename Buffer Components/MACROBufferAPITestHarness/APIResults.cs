using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft;
using System.IO;
using System.Xml.Xsl;
using System.Xml;
using System.Xml.XPath;

namespace MACROBufferAPITestHarness
{
    public partial class APIResults : Form
    {
        string _xml;

        public APIResults()
        {
            InitializeComponent();
            _xml = "";
        }

        public APIResults(string xml)
        {
            InitializeComponent();
            _xml = xml;
        }

        private void APIResults_Load(object sender, EventArgs e)
        {
            //Uri uri = new Uri("file:///C:/temp/results.xml");
            //webBrowser1.Url = uri;
            // read xsl file
            TextReader tr = new StreamReader(Application.StartupPath + "/xml-pretty-print.xsl");
            string xsl = tr.ReadToEnd();
            tr.Close();

            string html = MakeHTML(_xml, xsl);

            // webbrowser display empty document
            axWebBrowser1.Navigate("about:blank");

            // show string xml
            // create an IHTMLDocument2
            mshtml.IHTMLDocument2 doc = (mshtml.IHTMLDocument2)axWebBrowser1.Document;

            // write to the doc
            doc.clear();
            doc.writeln(html);
            //doc.writeln("<h1>hello</h1><br><br><h4>hello</h4>");

            doc.close();
        }

        public string MakeHTML(string xml, string xsl)
        {
            // Load the XML string into an XPathDocument.

            StringReader xmlStringReader = new StringReader(xml);
            XPathDocument xPathDocument = new XPathDocument(xmlStringReader);

            // Create a reader to read the XSL.

            StringReader xslStringReader = new StringReader(xsl);
            XmlTextReader xslTextReader = new XmlTextReader(xslStringReader);

            // Load the XSL into an XslTransform.

            XslTransform xslTransform = new XslTransform();
            xslTransform.Load(xslTextReader, null, GetType().Assembly.Evidence);

            // Perform the actual transformation and output an HTML string.

            StringWriter htmlStringWriter = new StringWriter();
            xslTransform.Transform(xPathDocument, null, htmlStringWriter, null);
            string html = htmlStringWriter.ToString();

            // Close all our readers and writers.
            xmlStringReader.Close();
            xslStringReader.Close();
            xslTextReader.Close();
            htmlStringWriter.Close();

            // Done, return the created HTML code.
            return html;
        }
    }
}