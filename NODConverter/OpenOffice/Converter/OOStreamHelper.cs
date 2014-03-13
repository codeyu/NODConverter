using System;
using System.Collections.Generic;
using System.Text;

namespace NODConverter.OpenOffice.Converter
{
    public class XOutputStreamWrapper : unoidl.com.sun.star.io.XOutputStream
    {
        private System.IO.Stream ostream;

        public XOutputStreamWrapper(System.IO.Stream ostream)
        {
            this.ostream = ostream;
        }

        public void closeOutput()
        {
            ostream.Close();
        }

        public void flush()
        {
            ostream.Flush();
        }

        public void writeBytes(byte[] b)
        {
            ostream.Write(b, 0, b.Length);
        }
    }

    public class XInputStreamWrapper : unoidl.com.sun.star.io.XInputStream, unoidl.com.sun.star.io.XSeekable
    {
        private System.IO.Stream instream;

        public XInputStreamWrapper(System.IO.Stream instream)
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
            instream.Seek(nToSkip, System.IO.SeekOrigin.Current);
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
            instream.Seek(pos, System.IO.SeekOrigin.Begin);
        }
    } 
}
