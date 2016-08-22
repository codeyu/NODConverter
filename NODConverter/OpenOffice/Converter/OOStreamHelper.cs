using System.IO;
using unoidl.com.sun.star.io;

namespace NODConverter.OpenOffice.Converter
{
    public class XOutputStreamWrapper : XOutputStream
    {
        private readonly Stream _ostream;

        public XOutputStreamWrapper(Stream ostream)
        {
            _ostream = ostream;
        }

        public void closeOutput()
        {
            _ostream.Close();
        }

        public void flush()
        {
            _ostream.Flush();
        }

        public void writeBytes(byte[] b)
        {
            _ostream.Write(b, 0, b.Length);
        }
    }

    public class XInputStreamWrapper : XInputStream, XSeekable
    {
        private Stream instream;

        public XInputStreamWrapper(Stream instream)
        {
            this.instream = instream;
        }

        // XInputStream interface 
        public int readBytes(out byte[] data, int length)
        {
            int remaining = (int)(instream.Length - instream.Position);
            int l = length > remaining ? remaining : length;
            data = new byte[length];
            return instream.Read(data, 0, l);
        }

        public int readSomeBytes(out byte[] data, int length)
        {
            int remaining = (int)(instream.Length - instream.Position);
            int l = length > remaining ? remaining : length;
            data = new byte[length];
            return instream.Read(data, 0, l);
        }

        public void skipBytes(int nToSkip)
        {
            instream.Seek(nToSkip, SeekOrigin.Current);
        }

        public int available()
        {
            return (int)(instream.Length - instream.Position);
        }

        public void closeInput()
        {
            instream.Close();
        }

        // XSeekable interface 
        public long getPosition()
        {
            return instream.Position;
        }

        public long getLength()
        {
            return instream.Length;
        }

        public void seek(long pos)
        {
            instream.Seek(pos, SeekOrigin.Begin);
        }
    } 
}
