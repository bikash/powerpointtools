using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using PowerPointLaTeX.Properties;

namespace PowerPointLaTeX
{
    class LaTeXServiceRegistry
    {
        private Dictionary<string, ILaTeXService> services = new Dictionary<string, ILaTeXService>();
        private string[] serviceNames;

        public LaTeXServiceRegistry() {
            Initialize();
        }
        
        private void Initialize() {
            // from: http://www.codeproject.com/KB/architecture/CSharpClassFactory.aspx
            // Get the assembly that contains this code

            Assembly asm = Assembly.GetCallingAssembly();
            
            // Get a list of all the types in the assembly

            Type[] allTypes = asm.GetTypes();
            foreach (Type type in allTypes)
            {
                // Only scan classes that arn't abstract
                if (type.IsClass && !type.IsAbstract && type.GetInterface("ILaTeXService") != null)
                {
                    ILaTeXService service = (ILaTeXService)asm.CreateInstance(type.FullName);
                    services.Add(service.SeriveName, service);
                }
            }

            var keys = services.Keys;
            serviceNames = new string[keys.Count];
            keys.CopyTo(serviceNames, 0);
        }

        public ILaTeXService GetService(string serviceName) {
            return services[serviceName];
        }

        public string[] ServiceNames {
            get {
                return serviceNames;
            }
        }

        public ILaTeXService Service {
            get {
                ILaTeXService service;
                if( !services.TryGetValue( Settings.Default.LatexService, out service ) ) {
                    var keyValuePair = services.First();
                    Settings.Default.LatexService = keyValuePair.Key;
                    service = keyValuePair.Value;
                }
                return service;
            }
        }
        
    }
}
