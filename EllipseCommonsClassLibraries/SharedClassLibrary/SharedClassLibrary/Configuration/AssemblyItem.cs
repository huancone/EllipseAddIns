using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.IO;
using System.Diagnostics;

//Shared Class Library - ExcelStyleCells
//Desarrollado por:
//Héctor J Hernández R <hernandezrhectorj@gmail.com>
//Hugo A Mendoza B <hugo.mendoza@hambings.com.co>

namespace SharedClassLibrary.Configuration
{
    public partial class Settings
    {
        public static Assembly GetLastAssembly()
        {
            var currentAssembly = Assembly.GetExecutingAssembly();
            var callerAssemblies = new StackTrace().GetFrames()
                .Select(x => x.GetMethod().ReflectedType.Assembly).Distinct()
                .Where(x => x.GetReferencedAssemblies().Any(y => y.FullName == currentAssembly.FullName));
            var initialAssembly = callerAssemblies.Last();
            
            return initialAssembly;
        }
        public class AssemblyItem
        {
            public string AssemblyTitle;
            public string AssemblyVersion;
            public string AssemblyDescription;
            public string AssemblyProduct;
            public string AssemblyCopyright;
            public string AssemblyCompany;
            public string AssemblyDeveloper1;
            public string AssemblyDeveloper2;

            private Assembly _addIAssembly;

            public AssemblyItem(Assembly assembly)
            {
                if (assembly == null)
                    return;
                _addIAssembly = assembly;
                AssemblyTitle = GetAssemblyTitle(_addIAssembly);
                AssemblyVersion = GetAssemblyVersion(_addIAssembly);
                AssemblyDescription = GetAssemblyDescription(_addIAssembly);
                AssemblyProduct = GetAssemblyProduct(_addIAssembly);
                AssemblyCopyright = GetAssemblyCopyright(_addIAssembly);
                AssemblyCompany = GetAssemblyCompany(_addIAssembly);
                AssemblyDeveloper1 = "hernandezrhectorj@gmail.com";
                AssemblyDeveloper2 = "huancone@gmail.com";
            }

            private string GetAssemblyTitle(Assembly assembly)
            {
                var attributes = assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false);

                if (attributes.Length <= 0)
                    return Path.GetFileNameWithoutExtension(assembly.CodeBase);
                var titleAttribute = (AssemblyTitleAttribute) attributes[0];

                return !string.IsNullOrWhiteSpace(titleAttribute.Title) ? titleAttribute.Title : Path.GetFileNameWithoutExtension(assembly.CodeBase);
            }

            public string GetAssemblyVersion(Assembly assembly)
            {
                return assembly.GetName().Version.ToString();
            }

            public string GetAssemblyDescription(Assembly assembly)
            {

                var attributes = assembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyDescriptionAttribute) attributes[0]).Description;

            }

            public string GetAssemblyProduct(Assembly assembly)
            {

                var attributes = assembly.GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyProductAttribute) attributes[0]).Product;

            }

            public string GetAssemblyCopyright(Assembly assembly)
            {

                var attributes = assembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyCopyrightAttribute) attributes[0]).Copyright;

            }

            public string GetAssemblyCompany(Assembly assembly)
            {

                var attributes = assembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0) return "";
                return ((AssemblyCompanyAttribute) attributes[0]).Company;

            }
        }
    }
}
