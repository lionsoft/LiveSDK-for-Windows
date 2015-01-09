using System.Collections.Generic;

namespace Microsoft.Live
{
    internal static class DictHelper
    {
        public static T Get<T>(this IDictionary<string, object> itemInfo, string name, T defValue = default(T))
        {
            if (itemInfo != null)
            {
                try
                {
                    object res;
                    if (itemInfo.TryGetValue(name, out res))
                        return (T)res;
                }
                catch
                {
                }
            }
            return defValue;
        }

        public static string Get(this IDictionary<string, object> itemInfo, string name)
        {
            return itemInfo.Get(name, "");
        }
    }
}