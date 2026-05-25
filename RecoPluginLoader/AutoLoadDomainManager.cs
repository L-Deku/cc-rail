using System;
using System.Globalization;
using System.IO;
using System.Reflection;

namespace RecoPluginLoader
{
    public sealed class AutoLoadDomainManager : AppDomainManager
    {
        public override void InitializeNewDomain(AppDomainSetup appDomainInfo)
        {
            base.InitializeNewDomain(appDomainInfo);
            LoadPlugin("RecoExpandPanel.dll", "RecoNet.FormPanel", "InstallOnIdle");
            LoadPlugin("RecoQuotaRecommend.dll", "RecoQuotaRecommend.QuotaRecommendPanel", "InstallOnIdle");
        }

        private static void LoadPlugin(string assemblyFileName, string typeName, string methodName)
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string assemblyPath = Path.Combine(baseDir, assemblyFileName);
                if (!File.Exists(assemblyPath))
                {
                    Log("Plugin not found: " + assemblyFileName);
                    return;
                }

                Assembly assembly = Assembly.LoadFrom(assemblyPath);
                Type type = assembly.GetType(typeName, false);
                if (type == null)
                {
                    Log("Plugin type not found: " + typeName);
                    return;
                }

                MethodInfo method = type.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static);
                if (method == null)
                {
                    Log("Plugin method not found: " + typeName + "." + methodName);
                    return;
                }

                method.Invoke(null, null);
                Log("Plugin loaded: " + assemblyFileName + " -> " + typeName + "." + methodName);
            }
            catch (Exception ex)
            {
                Log("Plugin load failed: " + assemblyFileName + " " + ex);
            }
        }

        private static void Log(string message)
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RecoPluginLoader.log");
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine);
            }
            catch
            {
            }
        }
    }
}
