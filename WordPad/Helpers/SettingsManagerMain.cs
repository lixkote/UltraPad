using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel.Resources.Core;
using Windows.Storage;

namespace WordPad.Helpers
{
    public class SettingsManagerMain
    {
        public object GetSetting(string key)
        {
            ApplicationDataContainer LocalSettings = ApplicationData.Current.LocalSettings;
            return LocalSettings.Values[key];
        }

        public string GetSettingString(string key)
        {
            ApplicationDataContainer LocalSettings = ApplicationData.Current.LocalSettings;
            return (string)LocalSettings.Values[key];
        }

        public void SetSetting(string key, object value)
        {
            ApplicationDataContainer LocalSettings = ApplicationData.Current.LocalSettings;
            LocalSettings.Values[key] = value;
        }

        public void InitializeDefaults()
        {
            /// This function is responsible for applying the default settings on the app's first startup.
            /// Modifying these will change the default settings of the application:     

            // Get the default resource context for the app
            ResourceContext defaultContext = ResourceContext.GetForCurrentView();
            var localSettings = ApplicationData.Current.LocalSettings;

            // Check if the app has been launched before and if not, set the default settings.
            if (localSettings.Values["FirstRun"] == null)
            {
                // Set the settings values to the default ones
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["unit"] = "inches";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["theme"] = "System";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["IsDarkThemeEditor"] = false;
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["fontfamily"] = "Calibri";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["fontsize"] = "11";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["papersize"] = "A4";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["papersource"] = "Auto";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["pagesetupBmargin"] = "0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["pagesetupRmargin"] = "0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["pagesetupTmargin"] = "0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["pagesetupLmargin"] = "0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["isprintpagenumbers"] = "no";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["indentationL"] = "0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["indentationR"] = "0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["indentationFL"] = "0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["is10ptenabled"] = "no";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["alignment"] = "Left";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["orientation"] = "Portrait";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["linespacing"] = "1,0";
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["textwrapping"] = "wrapruler";

                // Set the value to indicate that the app has been launched
                localSettings.Values["FirstRun"] = false;
            }
        }
    }
}
