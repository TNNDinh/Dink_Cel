using System;
using System.IO;
using System.Reflection;

namespace DinkCel
{
    internal static class EmbeddedDependencies
    {
        private static bool installed;

        public static void Install()
        {
            if (installed) return;
            installed = true;
            AppDomain.CurrentDomain.AssemblyResolve += delegate(object sender, ResolveEventArgs args)
            {
                string resource = new AssemblyName(args.Name).Name + ".dll";
                Assembly owner = Assembly.GetExecutingAssembly();
                using (Stream stream = owner.GetManifestResourceStream(resource))
                {
                    if (stream == null) return null;
                    var bytes = new byte[stream.Length];
                    int offset = 0;
                    while (offset < bytes.Length)
                    {
                        int read = stream.Read(bytes, offset, bytes.Length - offset);
                        if (read <= 0) throw new EndOfStreamException(resource);
                        offset += read;
                    }
                    return Assembly.Load(bytes);
                }
            };
        }
    }
}
