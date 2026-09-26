using System;
using System.Resources;
using System.Collections;
using System.Collections.Generic;
namespace DWG2PDF
{
    public static class ResxFileHelper
    {
        public static void WriteToResxFile(string resxFilePath, Dictionary<string, string> keyValuePairs)
        {
            using (ResXResourceWriter resxWriter = new ResXResourceWriter(resxFilePath))
            {
                foreach (var kvp in keyValuePairs)
                {
                    resxWriter.AddResource(kvp.Key, kvp.Value);
                }

                resxWriter.Generate();
            }
        }

        public static string ReadFromResxFile(string resxFilePath, string key)
        {
            using (ResXResourceReader resxReader = new ResXResourceReader(resxFilePath))
            {
                var resourceSet = resxReader.GetEnumerator();
                while (resourceSet.MoveNext())
                {
                    var entry = (DictionaryEntry)resourceSet.Entry;
                    if (entry.Key.ToString() == key)
                    {
                        return entry.Value.ToString();
                    }
                }
            }

            return null;
        }
    }
}