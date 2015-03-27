using Hyland.Unity;

namespace REUnityLibrary
{
    public class Utility
    {
        private Hyland.Unity.Application _app;
        private string _diagnosticLevel_ConfigurationItemName = "";

        public Utility(Hyland.Unity.Application app)
        {
            _app = app;
            _diagnosticLevel_ConfigurationItemName = "DiagnosticLevel";
        }

        public Utility(Hyland.Unity.Application app, string diagnosticLevel_ConfigurationItemName)
        {
            _app = app;
            _diagnosticLevel_ConfigurationItemName = diagnosticLevel_ConfigurationItemName;
        }

        #region SetDiagnosticLevel
        public void SetDiagnosticLevel()
        {
            string diagnosticLevel = "ERROR";
            bool success = _app.Configuration.TryGetValue(_diagnosticLevel_ConfigurationItemName, out diagnosticLevel);

            switch (diagnosticLevel.ToUpper())
            {
                case "INFO":
                    _app.Diagnostics.Level = Diagnostics.DiagnosticsLevel.Info;
                    break;
                case "WARNING":
                    _app.Diagnostics.Level = Diagnostics.DiagnosticsLevel.Warning;
                    break;
                case "VERBOSE":
                    _app.Diagnostics.Level = Diagnostics.DiagnosticsLevel.Verbose;
                    break;
                default:
                    _app.Diagnostics.Level = Diagnostics.DiagnosticsLevel.Error;
                    break;
            }

        }
        #endregion
    }
}
