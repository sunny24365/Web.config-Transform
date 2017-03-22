using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Web.config_Transform
{
    internal class TransformHelper
    {
        private bool _tranfmAppSettings;
        private bool _tranfmConnSettings;
        private bool _tranfmEndpoints;
        private string _excludedEndpoints;

        private string[] ExcludedEndpoints
        {
            get 
            {
                var endpoints = new string[] { };
                if (!string.IsNullOrEmpty(_excludedEndpoints))
                    endpoints = _excludedEndpoints.Split(',');
                
                return endpoints;
            }
        }

        public TransformHelper(bool tranfmAppSettings, bool tranfmConnSettings, bool tranfmEndpoints, string excludedEndpoints)
        {
            _tranfmAppSettings = tranfmAppSettings;
            _tranfmConnSettings = tranfmConnSettings;
            _tranfmEndpoints = tranfmEndpoints;
            _excludedEndpoints = excludedEndpoints;
        }

        public static DTE2 GetActiveIDE()
        {
            // Get an instance of currently running Visual Studio IDE.
            DTE2 dte2 = Package.GetGlobalService(typeof(DTE)) as DTE2;
            return dte2;
        }

        public static IList<Project> Projects()
        {
            Projects projects = GetActiveIDE().Solution.Projects;
            List<Project> list = new List<Project>();
            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                var project = item.Current as Project;
                if (project == null)
                {
                    continue;
                }

                if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(project));
                }
                else
                {
                    list.Add(project);
                }
            }

            return list;
        }

        private static IEnumerable<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            List<Project> list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                {
                    continue;
                }

                // If this is another solution folder, do a recursive call, otherwise add
                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(subProject));
                }
                else
                {
                    list.Add(subProject);
                }
            }
            return list;
        }

        private bool IsHavingWebConfig(Project project)
        {
            bool isHavingWebConfig = false;
            for (var i = 1; i <= project.ProjectItems.Count; i++)
            {
                if (project.ProjectItems.Item(i).Name.Equals("Web.config", System.StringComparison.CurrentCultureIgnoreCase))
                {
                    isHavingWebConfig = true;
                    break;
                }
            }

            return isHavingWebConfig;
        }

        private Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane GetOutputWindow()
        {
            var outWindow = Package.GetGlobalService(typeof(Microsoft.VisualStudio.Shell.Interop.SVsOutputWindow)) as Microsoft.VisualStudio.Shell.Interop.IVsOutputWindow;

            System.Guid generalPaneGuid = Microsoft.VisualStudio.VSConstants.GUID_OutWindowDebugPane; // P.S. There's also the GUID_OutWindowDebugPane available.

            Microsoft.VisualStudio.Shell.Interop.IVsOutputWindowPane generalPane;

            outWindow.GetPane(ref generalPaneGuid, out generalPane);

            return generalPane;
        }

        internal void Transform()
        {
            var outputWindow = GetOutputWindow();
            if (outputWindow != null)
            {
                outputWindow.Activate(); // Brings this pane into view
                outputWindow.OutputString("==================== Web.config Transform - STARTED ====================" + System.Environment.NewLine);
            }

            foreach (var proj in Projects())
            {
                if (IsHavingWebConfig(proj))
                {
                    string projDir = System.IO.Path.GetDirectoryName(proj.FullName);

                    //Apply Tranformation
                    UpdateConfig(projDir, proj.ConfigurationManager.ActiveConfiguration.ConfigurationName);

                    if (outputWindow != null) outputWindow.OutputString(" Transformed - " + projDir + System.Environment.NewLine);
                }
            }

            if (outputWindow != null) outputWindow.OutputString("==================== Web.config Transform - COMPLETED ====================" + System.Environment.NewLine);
        }

        /// <summary>
        /// Transforms web.config based on current active configuration
        /// </summary>
        /// <param name="projDir">Project that contains web.config</param>
        /// <param name="configName">Current Active Configuration selected in Configuration Manager</param>
        internal void UpdateConfig(string projDir, string configName)
        {
            string transformFilePath = string.Format(@"{0}\Web.{1}.config", projDir, configName);
            string configFilePath = string.Format(@"{0}\Web.config", projDir);

            if (System.IO.File.Exists(transformFilePath) && System.IO.File.Exists(configFilePath))
            {
                #region Transformation Region
                
                //Get - Transformation file
                var transformFileDoc = XDocument.Load(transformFilePath);

                //Get - Configuration file
                var configFileDoc = XDocument.Load(configFilePath);

                //Get all elements of transformation file
                var transformFileElements = transformFileDoc.Elements();
                if (transformFileElements.Any())
                {
                    //Get all elements of config file
                    var configFileElements = configFileDoc.Elements();

                    #region AppSettings

                    if (_tranfmAppSettings)
                    {
                        var transformFileAppSettings = transformFileElements.Descendants("appSettings").Elements().Where(elem => !elem.NodeType.Equals(XmlNodeType.Comment));
                        if (transformFileAppSettings.Any())
                        {
                            var configFileAppSettings = configFileElements.Descendants("appSettings").Elements().Where(elem => !elem.NodeType.Equals(XmlNodeType.Comment));

                            UpdateElement(transformFileAppSettings, configFileAppSettings, "key", "value");

                        }
                    }

                    #endregion

                    #region Connection Strings

                    if (_tranfmConnSettings)
                    {
                        var transformFileConnStrings = transformFileElements.Descendants("connectionStrings").Elements().Where(elem => !elem.NodeType.Equals(XmlNodeType.Comment));
                        if (transformFileConnStrings.Any())
                        {
                            var configFileConnStrings = configFileElements.Descendants("connectionStrings").Elements().Where(elem => !elem.NodeType.Equals(XmlNodeType.Comment));

                            UpdateElement(transformFileConnStrings, configFileConnStrings, "name", "connectionString");
                        } 
                    }

                    #endregion

                    #region End points

                    if (_tranfmEndpoints)
                    {
                        var transformFileEndpoints = transformFileElements.Descendants("system.serviceModel").Descendants("client").Elements().Where(elem => !elem.NodeType.Equals(XmlNodeType.Comment));

                        if (transformFileEndpoints.Any())
                        {
                            //Filter endpoints based on Options provided
                            var validtransformFileEndpoints = new List<XElement>();
                            foreach (var endpoint in transformFileEndpoints)
                            {
                                if (!(ExcludedEndpoints.Any(a => a.Equals(endpoint.Attribute("name").Value) || endpoint.Attribute("address").Value.Contains(a.Replace("*", "")))))
                                {
                                    validtransformFileEndpoints.Add(endpoint);
                                }
                            }

                            var configFileEndpoints = configFileElements.Descendants("system.serviceModel").Descendants("client").Elements().Where(elem => !elem.NodeType.Equals(XmlNodeType.Comment));

                            UpdateElement(validtransformFileEndpoints, configFileEndpoints, "name", "address");
                        } 
                    }

                    #endregion

                    //Save destination file
                    configFileDoc.Save(configFilePath, SaveOptions.None);
                }

                #endregion

            }

        }

        private static void UpdateElement(IEnumerable<XElement> transformFileSettings, IEnumerable<XElement> configFileSettings, 
            string attrToFind, string attrToReplace)
        {
            //Get AppSetting Keys
            var tranFileAppSettingsKeys = transformFileSettings.Select(d => d.Attribute(attrToFind).Value);
            var configFileAppSettingsKeys = configFileSettings.Select(d => d.Attribute(attrToFind).Value);

            //Loop through Keys and transform Values
            foreach (var key in tranFileAppSettingsKeys.Where(i => configFileAppSettingsKeys.Contains(i)))
            {
                var tranFileElement = transformFileSettings.Where(i => i.Attribute(attrToFind).Value.Equals(key)).FirstOrDefault();
                var configFileElement = configFileSettings.Where(i => i.Attribute(attrToFind).Value.Equals(key)).FirstOrDefault();

                var value = tranFileElement.Attribute(attrToReplace).Value;
                var configFileAttribute = configFileElement.Attribute(attrToReplace);

                if (!configFileAttribute.Value.Equals(value))
                    configFileAttribute.SetValue(value);
            }
        }
        
    }
}
