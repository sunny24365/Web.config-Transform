namespace Web.config_Transform
{
    public class OptionPageGrid : Microsoft.VisualStudio.Shell.DialogPage
    {
        private bool _appSettings = true;

        [System.ComponentModel.Category("Web.config Transform")]
        [System.ComponentModel.DisplayName("Transform Appsettings")]
        [System.ComponentModel.Description("Transform Appsettings")]
        public bool AppSettings
        {
            get { return _appSettings; }
            set { _appSettings = value; }
        }

        private bool _connectionStrings = true;

        [System.ComponentModel.Category("Web.config Transform")]
        [System.ComponentModel.DisplayName("Transform Connection Strings")]
        [System.ComponentModel.Description("Transform Connection Strings")]
        public bool ConnectionStrings
        {
            get { return _connectionStrings; }
            set { _connectionStrings = value; }
        }

        private bool _endpoints = true;
        
        [System.ComponentModel.Category("Web.config Transform")]
        [System.ComponentModel.DisplayName("Transform Endpoints")]
        [System.ComponentModel.Description("Transform Endpoints")]
        public bool Endpoints
        {
            get { return _endpoints; }
            set { _endpoints = value; }
        }

        private string _excludedEndpoints = string.Empty;

        [System.ComponentModel.Category("Web.config Transform")]
        [System.ComponentModel.DisplayName("Excluded Endpoints")]
        [System.ComponentModel.Description("Excluded Endpoints")]
        public string ExcludedEndpoints
        {
            get { return _excludedEndpoints; }
            set { _excludedEndpoints = value; }
        }
    }
}
