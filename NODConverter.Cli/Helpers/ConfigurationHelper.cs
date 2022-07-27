using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;


public class ConfigurationHelper
{
    #region "Public Functions"

    public static string GetString(string pKey) { return ConfigurationManager.AppSettings[pKey]; }

    public static int GetInt32(string pKey) { return int.Parse(GetString(pKey)); }

    public static int[] GetInt32Array(string pKey) { return string.IsNullOrEmpty(GetString(pKey)) ? new int[] { } : GetString(pKey).Split(',').Select(int.Parse).ToArray(); }

    public static bool GetBool(string pKey) { return bool.Parse(GetString(pKey)); }

    public static DateTime GetDateTime(string pKey) { return DateTime.Parse(GetString(pKey)); }

    public static TimeSpan GetTimeSpan(string pKey)
    {
        string[] _item = GetString(pKey).Split(':');
        return new TimeSpan(int.Parse(_item[0]), int.Parse(_item[1]), int.Parse(_item[2]));
    }

    public static Guid GetGuid(string pKey) { return Guid.Parse(GetString(pKey)); }
    #endregion
}
