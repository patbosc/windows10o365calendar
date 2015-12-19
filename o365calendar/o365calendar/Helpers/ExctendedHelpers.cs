using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage.Streams;

namespace o365calendar.Helpers
{
    public static class Extensions
    {
        public static void WriteStringWithLength(this DataWriter w, string s)
        {
            w.WriteUInt32((uint)s.Length);
            w.WriteString(s);
        }

        public static string ReadString(this DataReader r)
        {
            return r.ReadString(r.ReadUInt32());
        }
    }
}
