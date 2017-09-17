using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TYWordCopy.Model;
using TYWordCopy.Util;
using Net.Sgoliver.NRtfTree.Core;
using Net.Sgoliver.NRtfTree.Util;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Text.RegularExpressions;
using MarkupConverter;
using System.Collections;

namespace TYWordCopy.Controller.Server
{

    abstract class IReplace
    {
        protected string type;
        protected string _content;
        protected Dictionary<string, object> attrs = new Dictionary<string, object>();

        public string content
        {
            get { return _content; }
            set { _content = value; }
        }

        public object attr(string key)
        {
            if (attrs.ContainsKey(key))
            {
                return attrs[key];
            }
            else
            {
                return null;
            }
        }

        public void attr(string key, object value)
        {
            if (attrs.ContainsKey(key))
            {
                attrs[key] = value;
            }
            else
            {
                attrs.Add(key, value);
            }
        }

        abstract public string toHtml();
    }

    class EquationObj : IReplace
    {
        public EquationObj()
        {
            this.type = "equation";
        }

        public override string toHtml()
        {
            if (content == null)
            {
                return "";
            }

            string width = attr("width").ToString();
            string height = attr("height").ToString();
            string latex = attr("latex").ToString();

            if (latex == "")
            {
                return "";
            }

            StringBuilder imgHtml = new StringBuilder();
            imgHtml.AppendFormat("<img src='data:image/gif;base64,{0}' class='kfformula' ", content);

            if (width != "" && height != "")
            {
                imgHtml.AppendFormat(" width='{0}' height='{1}' ", width, height);
            }

            imgHtml.AppendFormat("  data-latex='{0}' />", latex);

            return imgHtml.ToString();
        }
    }

    class pictObj : IReplace
    {

        public ImageFormat format;

        public pictObj()
        {
            this.type = "pict";
        }

        public override string toHtml()
        {
            if (content == null)
            {
                return "";
            }

            string width = attr("width").ToString();
            string height = attr("height").ToString();

            string dataformat;
            if (format == ImageFormat.Jpeg)
            {
                dataformat = "image/jpeg";
            }
            else if (format == ImageFormat.Png)
            {
                dataformat = "image/png";
            }
            else if (format == ImageFormat.Bmp)
            {
                dataformat = "image/bmp";
            }
            else
            {
                dataformat = "image/jpeg";
            }

            StringBuilder imgHtml = new StringBuilder();
            imgHtml.AppendFormat("<img src='data:{0};base64,{1}' ", dataformat, content);

            if (width != "" && height != "")
            {
                imgHtml.AppendFormat(" width='{0}' height='{1}'", width, height);
            }

            imgHtml.Append(" />");

            return imgHtml.ToString();
        }
    }

    class DataConvert
    {

        string rtfText;
        const string REPLACETAG = "ReplaceNode";

        ArrayList rtf_replace = new ArrayList();
        IDataObject iData;

        private TYWordCopyAppController _controller;

        public event EventHandler<DataConvertEventArgs> DataConverted;

        public class DataConvertEventArgs : EventArgs
        {
            public string Html;
            public string Rtf;

            public DataConvertEventArgs(string rtf, string html)
            {
                Html = html;
                Rtf = rtf;
            }
        }

        public DataConvert(TYWordCopyAppController controller)
        {
            _controller = controller;

            controller.ClipboardChanged += ConvertData;
            //controller.FocusChanged += ConvertData;
            controller.Errored += controller_Errored;
        }

        protected int iFileNum = 0;

        public string GetOutputFile(string strExt)
        {
            string strRet = Utils.GetTempPath();
            string strFileName;
            strFileName = string.Format("Output{0}.{1}", iFileNum++, strExt);
            return System.IO.Path.Combine(strRet, strFileName);
        }


        private static byte GetHexValue(char c)
        {
            if (c >= '0' && c <= '9')
                return (byte)(c - '0');

            if (c >= 'a' && c <= 'f')
                return (byte)(10 + (c - 'a'));

            if (c >= 'A' && c <= 'F')
                return (byte)(10 + (c - 'A'));

            throw new ArgumentException(null, "c");
        }

        private byte[] getOleBinByte(string hexData)
        {
            int start = hexData.IndexOf("d0cf11e0");
            if (start < 0)
                throw new ArgumentException(null, "Text does not contain a doc file.");

            int end = hexData.Length;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                bool highByte = true;
                byte b = 0;
                for (int i = start; i < end; i++)
                {
                    char c = hexData[i];
                    if (char.IsWhiteSpace(c))
                        continue;

                    if (highByte)
                    {
                        b = (byte)(16 * GetHexValue(c));
                    }
                    else
                    {
                        b |= GetHexValue(c);
                        memoryStream.WriteByte(b);
                    }
                    highByte = !highByte;
                }
                return memoryStream.ToArray();
            }
        }


        private void reset()
        {
            rtfText = null;
            rtf_replace.Clear();
        }

        private RtfTreeNode getReplaceNode()
        {
            int replaceID = rtf_replace.Count;
            RtfTreeNode textNode = new RtfTreeNode(RtfNodeType.Text, "["+ REPLACETAG + "]" + replaceID + "[/" + REPLACETAG + "]", false, 0);
            return textNode;
        }

        public string ImageToBase64(Image image, ImageFormat format)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                // Convert Image to byte[]
                image.Save(ms, format);
                byte[] imageBytes = ms.ToArray();

                // Convert byte[] to Base64 String
                string base64String = Convert.ToBase64String(imageBytes);
                return base64String;
            }
        }

        private EquationObj Wmf2Equation(ImageNode imageNode)
        {
            EquationObj equationObj = new EquationObj();
            ConvertEquation ce = new ConvertEquation();

            EquationInput ei = new EquationInputFileWMF(imageNode.GetByteData());
            string strEquation = ce.ConvertToText(ei);
            Regex regex = new Regex(@"<([^>]+)>(.*?)</\1>");
            var v = regex.Match(strEquation);
            string latexString = v.Groups[2].ToString();

            if (latexString.Length > 0)
            {
                equationObj.attr("latex", latexString);
            }
            
            string gifBase64Code = "";
            Image gifbit = ce.ConvertToGif(ei);
            
            if (gifbit != null)
            {
                gifBase64Code = ImageToBase64(gifbit, ImageFormat.Gif);

                equationObj.content = gifBase64Code;
                equationObj.attr("width", gifbit.Width);
                equationObj.attr("height", gifbit.Height);
            }
            
            return equationObj;
        }

        private ImageFormat getImageFormat(Image image)
        {
            if (ImageFormat.Jpeg.Equals(image.RawFormat))
            {
                return ImageFormat.Jpeg;
            }
            else if (ImageFormat.Png.Equals(image.RawFormat))
            {
                return ImageFormat.Png;
            }
            else if (ImageFormat.Gif.Equals(image.RawFormat))
            {
                return ImageFormat.Gif;
            }
            else
            {
                return ImageFormat.Jpeg;
            }
        }

        private RtfTree _ConvertMMath(RtfTree tree)
        {
            RtfNodeCollection nodes = tree.RootNode.SelectNodes("mmath");

            //RtfNodeCollection nodes = tree.RootNode.ChildNodes;

            foreach (RtfTreeNode n in nodes)
            {
                Console.WriteLine(n.ParentNode.Rtf);
                RtfTreeNode auxNode = null;
                if ((auxNode = n.ParentNode.SelectSingleNode("nonshppict")) != null)
                {

                    ImageNode imageNode = new ImageNode(auxNode.ParentNode.SelectSingleNode("pict").ParentNode);

                    ConvertEquation ce = new ConvertEquation();

                    EquationInput ei = new EquationInputFileEPS(imageNode.GetByteData());
                    EquationOutput eo = new EquationOutputFileGIF(GetOutputFile("gif"));
                    ce.Convert(ei,eo);
                    //string strEquation = ce.ConvertToText(ei);
                    //EquationObj equationObj = Wmf2Equation(imageNode);


                    //tree.RootNode.FirstChild.ReplaceChildDeep(n.ParentNode, getReplaceNode());

                    //rtf_replace.Add(equationObj);
                }
                else
                {
                    MessageBox.Show("'\result' node contains no images!");
                }


                tree.RootNode.FirstChild.RemoveChildDeep(n.ParentNode);
            }
            return tree;
        }

        private RtfTree _ConvertObject(RtfTree tree)
        {
            RtfNodeCollection nodes = tree.RootNode.SelectNodes("object");

            //RtfNodeCollection nodes = tree.RootNode.ChildNodes;

            foreach (RtfTreeNode n in nodes)
            {

                ObjectNode objectNode = new ObjectNode(n.ParentNode);
                
                if (objectNode.ObjectType == "objemb" && objectNode.ObjectClass == "Equation.DSMT4")
                {
                    RtfTreeNode resultNode = objectNode.ResultNode;

                    RtfTreeNode auxNode = null;

                    if ((auxNode = resultNode.SelectSingleNode("pict")) != null)
                    {

                        ImageNode imageNode = new ImageNode(auxNode.ParentNode);
                        EquationObj equationObj = Wmf2Equation(imageNode);


                        tree.RootNode.FirstChild.ReplaceChildDeep(n.ParentNode, getReplaceNode());

                        rtf_replace.Add(equationObj);
                    }
                    else
                    {
                        MessageBox.Show("'\result' node contains no images!");
                    }
                }

                tree.RootNode.FirstChild.RemoveChildDeep(n.ParentNode);
            }
            return tree;
        }

        private RtfTree _ConvertPict(RtfTree tree)
        {
            RtfNodeCollection shppict_nodes = tree.RootNode.SelectNodes("shppict");
            foreach (RtfTreeNode n in shppict_nodes)
            {
                tree.RootNode.FirstChild.RemoveChildDeep(n.ParentNode);
            }

            RtfNodeCollection pict_nodes = tree.RootNode.SelectNodes("pict");
            foreach (RtfTreeNode n in pict_nodes)
            {
                ImageNode imageNode = new ImageNode(n.ParentNode);

                pictObj pictObj = new pictObj();

                ImageConverter ic = new ImageConverter();
                Image img = (Image)ic.ConvertFrom(imageNode.GetByteData());
                ImageFormat format = getImageFormat(img);
                //string imgBase64Code = Convert.ToBase64String(imageNode.GetByteData());
                string imgBase64Code = ImageToBase64(img, format);

                pictObj.format = format;
                pictObj.content = imgBase64Code;
                pictObj.attr("width", imageNode.Width);
                pictObj.attr("height", imageNode.Height);

                tree.RootNode.FirstChild.ReplaceChildDeep(n.ParentNode.ParentNode, getReplaceNode());

                rtf_replace.Add(pictObj);

                tree.RootNode.FirstChild.RemoveChildDeep(n.ParentNode.ParentNode);
            }
            return tree;
        }

        private void ConvertData(object sender, ClipboardMonitor.ClipboardChangedEventArgs e)
        //private void ConvertData(object sender, System.Windows.Automation.AutomationFocusChangedEventArgs e)
        {
            //Console.WriteLine("ConvertData");
            iData = Clipboard.GetDataObject();

            if (!iData.GetDataPresent(DataFormats.Rtf))
            {
                return;
            }

            reset();

            rtfText = iData.GetData(DataFormats.Rtf).ToString();
            //Console.WriteLine(rtfText);


            RtfTree tree = new RtfTree();
            tree.LoadRtfText(rtfText);

            //tree = _ConvertMMath(tree);

            tree = _ConvertObject(tree);

            tree = _ConvertPict(tree);


            //Console.WriteLine(tree.Rtf);

            string html = RtfToHtmlConverter.ConvertRtfToHtml(tree.Rtf);
            //Console.WriteLine(html);

            string regularExpressionPattern = string.Format(@"\[{0}\](.*?)\[\/{1}\]", REPLACETAG, REPLACETAG);
            Regex regex = new Regex(regularExpressionPattern, RegexOptions.Singleline);
            MatchCollection collection = regex.Matches(html);

            foreach (Match m in collection)
            {
                string Matchstr = m.Groups[0].Value;
                string Rid = m.Groups[1].Value;
                int index = int.Parse(Rid);

                if (index < rtf_replace.Count)
                {
                    IReplace equationObj = (IReplace)rtf_replace[index];
                    string s = equationObj.toHtml();
                    //Console.WriteLine(m.Groups[0].Value + "\n\r");
                    //Console.WriteLine(s + "\n\r");
                    html = html.Replace(Matchstr, s);
                }
            }

            //Console.WriteLine(html + "\n\r");

            if (DataConverted != null)
            {
                DataConverted(this,new DataConvertEventArgs(rtfText,html));
            }
        }


        void controller_Errored(object sender, ErrorEventArgs e)
        {
            MessageBox.Show(e.GetException().ToString(), String.Format(I18N.GetString("Error: {0}"), e.GetException().Message));
            Console.Write(e.GetException().ToString(), String.Format(I18N.GetString("Error: {0}"), e.GetException().Message));
        }

    }
}