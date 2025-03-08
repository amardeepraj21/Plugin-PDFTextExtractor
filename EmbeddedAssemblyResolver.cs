using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using System.IO;

namespace PDFTextExtractor
{
    public static class EmbeddedAssemblyResolver
    {
        public static void AttachResolver()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string resourceName = $"{Assembly.GetExecutingAssembly().GetName().Name}.{new AssemblyName(args.Name).Name}.dll";
                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        byte[] assemblyData = new byte[stream.Length];
                        stream.Read(assemblyData, 0, assemblyData.Length);
                        return Assembly.Load(assemblyData);
                    }
                }
                return null; // Return null if the assembly is not found
            };
        }
    }
}
