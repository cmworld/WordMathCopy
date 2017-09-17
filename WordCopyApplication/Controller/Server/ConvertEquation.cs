using System;
using System.Text;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using TYWordCopy.Util;
using MTSDKDN;

namespace TYWordCopy.Controller.Server
{

    class MTSDK
    {

        // c-tor
        public MTSDK() { }

        // vars
        protected bool m_bDidInit = false;

        // init
        public bool Init()
        {
            if (!m_bDidInit)
            {
                Int32 result = MathTypeSDK.Instance.MTAPIConnectMgn(MTApiStartValues.mtinitLAUNCH_AS_NEEDED, 30);
                if (result == MathTypeReturnValue.mtOK)
                {
                    m_bDidInit = true;
                    return true;
                }
                else
                    return false;
            }
            return true;
        }

        // de-init
        public bool DeInit()
        {
            if (m_bDidInit)
            {
                m_bDidInit = false;
                MathTypeSDK.Instance.MTAPIDisconnectMgn();
            }
            return true;
        }

    }

    #region EquationInput Classes
    abstract class EquationInput
    {
        // c-tor
        public EquationInput(string strInTrans)
        {
            if (!string.IsNullOrEmpty(strInTrans))
                this.strInTrans = strInTrans;
            else
                this.strInTrans = string.Empty;
        }

        protected short m_iType;
        public short iType
        {
            get { return m_iType; }
            protected set { m_iType = value; }
        }

        protected short m_iFormat;
        public short iFormat
        {
            get { return m_iFormat; }
            protected set { m_iFormat = value; }
        }

        // the equation as a string
        protected string m_strEquation;
        public string strEquation
        {
            get { return m_strEquation; }
            set { m_strEquation = value; }
        }

        // the equation as a byte array
        protected byte[] m_bEquation;
        public byte[] bEquation
        {
            get { return m_bEquation; }
            set { m_bEquation = value; }
        }

        // MTEF byte array
        protected byte[] m_bMTEF;
        public byte[] bMTEF
        {
            get { return m_bMTEF; }
            set { m_bMTEF = value; }
        }

        // MTEF byte array length
        protected int m_iMTEF_Length;
        public int iMTEF_Length
        {
            get { return m_iMTEF_Length; }
            set { m_iMTEF_Length = value; }
        }

        // MTEF string
        protected string m_strMTEF;
        public string strMTEF
        {
            get { return m_strMTEF; }
            set { m_strMTEF = value; }
        }

        // input translator
        protected string m_strInTrans;
        public string strInTrans
        {
            get { return m_strInTrans; }
            set { m_strInTrans = value; }
        }

        // the source equation file
        protected string m_strFileName;
        public string strFileName
        {
            get { return m_strFileName; }
            set { m_strFileName = value; }
        }

        // the source equation file
        protected byte[] m_strFileByte;
        public byte[] strFileByte
        {
            get { return m_strFileByte; }
            set { m_strFileByte = value; }
        }

        protected MTSDK sdk = new MTSDK();

        // get the equation from the source
        abstract public bool Get();

        // get binary MTEF
        abstract public bool GetMTEF();
    }


    abstract class EquationInputFile : EquationInput
    {
        public EquationInputFile(string strFileName, string strInTrans) : base(strInTrans)
        {
            this.strFileName = strFileName;
            iType = MTXFormEqn.mtxfmLOCAL;
        }

        public EquationInputFile(byte[] strFileByte, string strInTrans) : base(strInTrans)
        {
            this.strFileByte = strFileByte;

            this.strFileName = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

            /*
            FileStream _FileStream = new FileStream(strFileName, FileMode.Create, FileAccess.Write);
            _FileStream.Write(strFileByte, 0, strFileByte.Length);
            _FileStream.Close();
            */

            File.WriteAllBytes(strFileName, strFileByte);

            iType = MTXFormEqn.mtxfmLOCAL;
        }
    }

    class EquationInputFileWMF : EquationInputFile
    {
        public EquationInputFileWMF(string strFileName) : base(strFileName, "")
        {
            iFormat = MTXFormEqn.mtxfmMTEF;
        }

        public EquationInputFileWMF(byte[] strFileByte) : base(strFileByte, "")
        {
            iFormat = MTXFormEqn.mtxfmMTEF;
        }

        public override bool Get() { return true; }

        public override string ToString() { return "WMF file"; }

        public override bool GetMTEF()
        {
            Play();
            if (!Succeeded())
                return false;
            return true;
        }

        protected class WmfForm : Form
        {
            public WmfForm() { }
        }
        protected WmfForm wf = new WmfForm();

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        protected struct wmfHeader
        {
            public Int16 iComment;
            public Int16 ix1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
            public string strSig;
            public Int16 iVer;
            public Int32 iTotalLen;
            public Int32 iDataLen;
        };
        protected wmfHeader m_wmfHeader;

        protected Metafile m_metafile;
        protected const string m_strSig = "AppsMFC";
        protected bool m_succeeded = false;

        protected void Play()
        {
   
           // try
           // {
                m_succeeded = false;
                Graphics.EnumerateMetafileProc metafileDelegate;
                Point destPoint;

                m_metafile = new Metafile(strFileName);

                if (!m_metafile.GetMetafileHeader().IsWmfPlaceable())
                {
                    m_metafile = RestorePlaceableHeader(m_metafile);
                }

                metafileDelegate = new Graphics.EnumerateMetafileProc(MetafileCallback);
                destPoint = new Point(20, 10);
                Graphics graphics = wf.CreateGraphics();
                graphics.EnumerateMetafile(m_metafile, destPoint, metafileDelegate);
           // }
           // catch (Exception e)
           // {
           //     Console.WriteLine(e.Message);
           // }
        }

        protected bool Succeeded() { return m_succeeded; }

        protected bool MetafileCallback(
            EmfPlusRecordType recordType,
            int flags,
            int dataSize,
            IntPtr data,
            PlayRecordCallback callbackData)
        {
            byte[] dataArray = null;

            if (data != IntPtr.Zero)
            {
                dataArray = new byte[dataSize];
                Marshal.Copy(data, dataArray, 0, dataSize);

                if (recordType == EmfPlusRecordType.WmfEscape && dataSize >= Marshal.SizeOf(m_wmfHeader) && !m_succeeded)
                {
                    m_wmfHeader = (wmfHeader)RawDeserialize(dataArray, 0, m_wmfHeader.GetType());
                    if (m_wmfHeader.strSig.Equals(m_strSig, StringComparison.CurrentCultureIgnoreCase))
                    {
                        System.Text.ASCIIEncoding enc = new System.Text.ASCIIEncoding();
                        string strCompanyInfo = enc.GetString(dataArray, Marshal.SizeOf(m_wmfHeader), m_wmfHeader.iDataLen);
                        int iNull = strCompanyInfo.IndexOf('\0');
                        if (iNull >= 0)
                        {
                            int mtefStart = Marshal.SizeOf(m_wmfHeader) + iNull + 1;
                            iMTEF_Length = m_wmfHeader.iDataLen;
                            bMTEF = new byte[iMTEF_Length];
                            Array.Copy(dataArray, mtefStart, bMTEF, 0, iMTEF_Length);
                            m_succeeded = true;
                        }
                    }
                }
            }

            m_metafile.PlayRecord(recordType, flags, dataSize, dataArray);

            return true;
        }

        protected object RawDeserialize(byte[] rawData, int position, Type anyType)
        {
            int rawsize = Marshal.SizeOf(anyType);
            if (rawsize > rawData.Length)
                return null;
            IntPtr buffer = Marshal.AllocHGlobal(rawsize);
            Marshal.Copy(rawData, position, buffer, rawsize);
            object retobj = Marshal.PtrToStructure(buffer, anyType);
            Marshal.FreeHGlobal(buffer);
            return retobj;
        }


        private Metafile RestorePlaceableHeader(Metafile Metafile)
        {
            byte[] bHeader = new byte[22];

                byte[] bBounds = GetBounds(Metafile);

                // Create header for placeable WMF
                WmfPlaceableFileHeader newhdr = new WmfPlaceableFileHeader();
                newhdr.BboxBottom = BitConverter.ToInt16(bBounds, 4); //1280;        
                newhdr.BboxLeft = 0;
                newhdr.BboxRight = BitConverter.ToInt16(bBounds, 0); //5056        
                newhdr.BboxTop = 0;
                newhdr.Hmf = 0;
                newhdr.Inch = 1400;

                // evaluate checksum
                int checksum = 0;
                checksum ^= (newhdr.Key & 0x0000FFFF);
                checksum ^= (int)(((ulong)newhdr.Key & 0xFFFF0000) >> 16);
                checksum ^= newhdr.Hmf;
                checksum ^= newhdr.BboxLeft;
                checksum ^= newhdr.BboxTop;
                checksum ^= newhdr.BboxRight;
                checksum ^= newhdr.BboxBottom;
                checksum ^= newhdr.Inch;
                checksum ^= (newhdr.Reserved & 0x0000FFFF);
                newhdr.Checksum = (short)checksum;

                MemoryStream msHeader = new MemoryStream();
                msHeader.Write(Reverse(newhdr.Key), 0, 4);
                msHeader.Write(Reverse(newhdr.Hmf), 0, 2);
                msHeader.Write(Reverse(newhdr.BboxLeft), 0, 2);
                msHeader.Write(Reverse(newhdr.BboxTop), 0, 2);
                msHeader.Write(Reverse(newhdr.BboxRight), 0, 2);
                msHeader.Write(Reverse(newhdr.BboxBottom), 0, 2);
                msHeader.Write(Reverse(newhdr.Inch), 0, 2);
                msHeader.Write(Reverse(newhdr.Reserved), 0, 4);
                msHeader.Write(Reverse(newhdr.Checksum), 0, 2);
                msHeader.Position = 0;
                int n = msHeader.Read(bHeader, 0, Convert.ToInt32(msHeader.Length));
                msHeader.Close();

            FileStream fileStream = File.Open(strFileName, FileMode.Open);
            fileStream.SetLength(0);
            fileStream.Write(bHeader, 0, bHeader.Length);
            fileStream.Write(strFileByte, 0, strFileByte.Length);
            fileStream.Close();
            return new Metafile(strFileName);
        }

        private static byte[] Reverse(object isrc)
        {
            byte[] src;
            if (isrc.GetType().ToString() == "System.Int32")
            {
                src = BitConverter.GetBytes((int)isrc);
            }
            else
            {
                src = BitConverter.GetBytes((short)isrc);
            }


            if (!BitConverter.IsLittleEndian)
            {

                byte[] res = new byte[src.Length];
                int i = src.Length - 1;
                foreach (byte b in src)
                {
                    res[i] = b;
                    i--;
                }
                return res;
            }
            return src;
        }

        [DllImport("Gdi32.dll", ExactSpelling = false, CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateEnhMetaFile(IntPtr hdcRef, string lpFilename, [In] ref RECT lpRect, string lpDescription);


        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public long Left;
            public long Top;
            public long Right;
            public long Bottom;
        }

        // restore bounds for 22-bytes header
        private static byte[] bBounds = new byte[8];
        private static byte[] GetBounds(Metafile mf)
        {

            RECT bounds;
            bounds.Left = 0;
            bounds.Top = 0;
            bounds.Right = 0;
            bounds.Bottom = 0;

            IntPtr hEmfDC = CreateEnhMetaFile(IntPtr.Zero, null, ref bounds, null);

            Graphics grfx = Graphics.FromHdc(hEmfDC);
            grfx.EnumerateMetafile(mf, new Point(0, 0), new Graphics.EnumerateMetafileProc(MetafileResetCallback));
            grfx.Dispose();
            return bBounds;
        }

        private static bool MetafileResetCallback(EmfPlusRecordType recordType, int flags, int dataSize, IntPtr data, PlayRecordCallback callbackData)
        {
            byte[] dataArray = null;
            if (data != IntPtr.Zero)
            {
                dataArray = new byte[dataSize];
                Marshal.Copy(data, dataArray, 0, dataSize);

                if (recordType == EmfPlusRecordType.EmfSetWindowExtEx)
                {
                    bBounds = dataArray;
                }
            }
            return true;
        }
    }
    

    class EquationInputFileEPS : EquationInputFile
    {
        public EquationInputFileEPS(string strFileName)
            : base(strFileName, "")
        {
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public EquationInputFileEPS(byte[] strFileByte) : base(strFileByte, "")
        {
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public override string ToString() { return "EPS file"; }

        public override bool Get()
        {
            try
            {
                strEquation = System.IO.File.ReadAllText(strFileName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public override bool GetMTEF()
        {
            /*
			extracting MTEF from EPS files is described in the MathType SDK documentation, by default installed here:
			file:///C:/Program%20Files/MathTypeSDK/SDK/docs/MTEFstorage.htm#Translator%20Output
			*/
            const string strSig1 = "MathType";
            const string strSig2 = "MTEF";
            int iSig1Start = 0;
            while ((iSig1Start = strEquation.IndexOf(strSig1, iSig1Start)) >= 0)
            {
                int iSig2Start = strEquation.IndexOf(strSig2, iSig1Start + 1);
                int iDelimStart = iSig1Start + strSig1.Length;
                int iDelimLen = iSig2Start - iDelimStart;
                if (iSig2Start < 0 || iDelimLen != 1)
                {
                    iSig1Start++;
                    continue;
                }
                string strDelim = strEquation.Substring(iDelimStart, iDelimLen);
                int id1 = strEquation.IndexOf(strDelim, iSig1Start);
                int id2 = strEquation.IndexOf(strDelim, id1 + 1);
                int id3 = strEquation.IndexOf(strDelim, id2 + 1);
                int id4 = strEquation.IndexOf(strDelim, id3 + 1);
                int id5 = strEquation.IndexOf(strDelim, id4 + 1);
                int id6 = strEquation.IndexOf(strDelim, id5 + 1);
                m_strMTEF = strEquation.Substring(iSig1Start, id6 - iSig1Start + 1);
                bMTEF = System.Text.Encoding.ASCII.GetBytes(m_strMTEF);
                iMTEF_Length = bMTEF.Length;
                return true;
            }
            return false;
        }
    }
    #endregion

    #region EquationOutput Classes
    abstract class EquationOutput
    {
        // c-tor
        public EquationOutput(string strOutTrans)
        {
            if (!string.IsNullOrEmpty(strOutTrans))
                this.strOutTrans = strOutTrans;
            else
                this.strOutTrans = string.Empty;
        }

        protected EquationOutput() { }

        // properties
        protected short m_iType;
        public short iType
        {
            get { return m_iType; }
            protected set { m_iType = value; }
        }

        protected short m_iFormat;
        public short iFormat
        {
            get { return m_iFormat; }
            protected set { m_iFormat = value; }
        }

        private string m_strFileName;
        public string strFileName
        {
            get { return m_strFileName; }
            set { m_strFileName = value; }
        }

        private string m_strEquation;
        public string strEquation
        {
            get { return m_strEquation; }
            set { m_strEquation = value; }
        }

        // output translator
        protected string m_strOutTrans;
        public string strOutTrans
        {
            get { return m_strOutTrans; }
            set { m_strOutTrans = value; }
        }

        // save equation to its destination
        abstract public bool Put();
    }

    abstract class EquationOutputFile : EquationOutput
    {
        public EquationOutputFile(string strFileName, string strOutTrans)
            : base(strOutTrans)
        {
            this.strFileName = strFileName;
            iType = MTXFormEqn.mtxfmFILE;
        }

        protected EquationOutputFile(string strFileName)
            : base()
        {
            this.strFileName = strFileName;
            iType = MTXFormEqn.mtxfmFILE;
        }

        public override bool Put() { return true; }
    }

    class EquationOutputFileText : EquationOutputFile
    {
        public EquationOutputFileText(string strFileName, string strOutTrans)
            : base(strFileName, strOutTrans)
        {
            iType = MTXFormEqn.mtxfmLOCAL; // override base class as the convert function cannot directly write text files
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public override bool Put()
        {
            try
            {
                FileStream stream = new FileStream(strFileName, FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter writer = new StreamWriter(stream);
                writer.WriteLine(strEquation);
                writer.Close();
                stream.Close();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public override string ToString() { return "Text file"; }
    }

    class EquationOutputFileWMF : EquationOutputFile
    {
        public EquationOutputFileWMF(string strFileName)
            : base(strFileName)
        {
            iFormat = MTXFormEqn.mtxfmPICT;
        }

        public override string ToString() { return "WMF file"; }
    }


    class EquationOutputFileGIF : EquationOutputFile
    {
        public EquationOutputFileGIF(string strFileName)
            : base(strFileName)
        {
            iFormat = MTXFormEqn.mtxfmGIF;
        }

        public override string ToString() { return "GIF file"; }
    }
    #endregion


    #region ConvertEquation Class
    class ConvertEquation
    {

            protected EquationInput m_ei;
            protected EquationOutput m_eo;
            protected MTSDK m_sdk = new MTSDK();

            // c-tor
            public ConvertEquation() { }

            // convert
            virtual public bool Convert(EquationInput ei, EquationOutput eo)
            {
                m_ei = ei;
                m_eo = eo;
                return Convert();
            }

            public string ConvertToText(EquationInput ei)
            {
                m_ei = ei;
                m_eo = new EquationOutputFileText("", "Texvc.tdl");

                if (m_ei.Get())
                {
                    // Console.WriteLine("Get MTEF");
                    if (m_ei.GetMTEF())
                    {
                        if (ConvertToOutput())
                        {
                            return m_eo.strEquation;
                        }
                    }
                }

                return "";
            }

            public Image ConvertToGif(EquationInput ei)
            {
                m_ei = ei;

                string fileName = Path.GetTempPath() + Guid.NewGuid().ToString() + ".tmp";
                m_eo = new EquationOutputFileGIF(fileName);

                if (m_ei.Get())
                {
                    // Console.WriteLine("Get MTEF");
                    if (m_ei.GetMTEF())
                    {
                        if (ConvertToOutput())
                        {
 
                            Image img = Bitmap.FromFile(fileName);
                            //File.Delete(fileName);

                            return img;
                        }
                    }
                }

                return null;
            }

        virtual protected bool Convert()
            {
                bool bReturn = false;

                Console.WriteLine("Converting {0} to {1}", m_ei.ToString(), m_eo.ToString());

                Console.WriteLine("Get equation: {0}", m_ei.strFileName);
                if (m_ei.Get())
                {
                    Console.WriteLine("Get MTEF");
                    if (m_ei.GetMTEF())
                    {
                        Console.WriteLine("Convert Equation");
                        if (ConvertToOutput())
                        {
                            Console.WriteLine("Write equation: {0}", m_eo.strFileName);
                            if (m_eo.Put())
                                bReturn = true;
                        }
                    }
                }

                Console.WriteLine("Convert success: {0}\r\n", bReturn.ToString());
                return bReturn;
            }

            protected bool SetTranslator()
            {
                if (string.IsNullOrEmpty(m_eo.strOutTrans))
                    return true;

                Int32 stat = MathTypeSDK.Instance.MTXFormSetTranslatorMgn(
                    MTXFormSetTranslator.mtxfmTRANSL_INC_NAME + MTXFormSetTranslator.mtxfmTRANSL_INC_DATA,
                    m_eo.strOutTrans);
                return stat == MathTypeReturnValue.mtOK;
            }

            protected bool ConvertToOutput()
            {
                bool bResult = false;

                //try
                //{
                    if (!m_sdk.Init())
                        return false;

                    if (MathTypeSDK.Instance.MTXFormResetMgn() == MathTypeReturnValue.mtOK &&
                        SetTranslator())
                    {
                        Int32 stat = 0;
                        Int32 iBufferLength = 5000;
                        StringBuilder strDest = new StringBuilder(iBufferLength);
                        MTAPI_DIMS dims = new MTAPI_DIMS();

                        // convert
                        stat = MathTypeSDK.Instance.MTXFormEqnMgn(
                            m_ei.iType,
                            m_ei.iFormat,
                            m_ei.bMTEF,
                            m_ei.iMTEF_Length,
                            m_eo.iType,
                            m_eo.iFormat,
                            strDest,
                            iBufferLength,
                            m_eo.strFileName,
                            ref dims);

                        // save equation
                        if (stat == MathTypeReturnValue.mtOK)
                        {
                            m_eo.strEquation = strDest.ToString();
                            bResult = true;
                        }
                    }

                    m_sdk.DeInit();
               // }
               // catch (Exception e)
               // {
               //     Console.WriteLine(e.Message);
               //     return false;
               // }

                return bResult;
            }
    }
        #endregion
}
