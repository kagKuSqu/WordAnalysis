using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Xml;

namespace WordAnalysis
{
    using System.Text.RegularExpressions;

    using Microsoft.Scripting.Utils;

    public static class XmlNodeListExtension
    {
        public static IEnumerable<XmlNode> ToEnumerable(this XmlNodeList source)
        {
            foreach (XmlNode node in source)
            {
                yield return node;
            }
        }
    }

    public static class MatchCollectionExtension
    {
        public static IEnumerable<string> Values(this MatchCollection source)
        {
            foreach (Match match in source)
            {
                yield return match.Value;
            }
        }
    }

    public static class ListExtension
    {
        public static T Get<T>(this List<T> source, int index=0) where T : class
        {
            if (source == null || source.Count <= index)
            {
                return null;
            }

            return source[index];
        }
        
    }
    public static class StringExtension
    {
        public static string IsNull(this string source, string @default = "")
        {
            return source.IsEmpty() ? @default : source;
        }

        public static IEnumerable<string> MatchesSplit(this string source, string pattern = "")
        {
            int[] arr = source.Matches(pattern).Select(time => source.IndexOf(time)).ToArray();
            for (int i = 0; i < arr.Length; i++)
            {
                string str;
                var current = arr[i];
                if (i+1==arr.Length)
                {
                    str = source.SubstringEx(current);
                }
                else
                {
                    var next = arr[i + 1];
                    var offset = next - current;
                    str = source.SubstringEx(current, offset);
                }
                yield return str;
            }
        }

        public static string SubstringEx(this string source,int start,int len=0)
        {
            if (start>source.Length||len>source.Length-start)
            {
                return string.Empty;
            }
            if (len!=0)
            {
                return source.Substring(start, len);
            }
            return source.Substring(start);
        }

        public static string ReplaceRegex(this string source, string pattern = "",string replacement="")
        {
            source.Matches(pattern).ForEach(m => source = source.Replace(m, replacement));
            return source;
        }
        public static List<string> Matches(this string source, string pattern = "")
        {
            if (source.IsEmpty())
            {
                return new List<string>();
            }
            return System.Text.RegularExpressions.Regex.Matches(source, pattern).Values().ToList();
        }
        public static string MatchesJoin(this string source, string pattern = "", string separator = "")
        {
            if (source.IsEmpty())
            {
                return string.Empty;
            }
            return System.Text.RegularExpressions.Regex.Matches(source, pattern).Values().ToList().JoinStrings(separator);
        }
        public static string MatchesJoinTrim(this string source, string pattern = "", string separator = "")
        {
            if (source.IsEmpty())
            {
                return string.Empty;
            }
            return System.Text.RegularExpressions.Regex.Matches(source, pattern).Values().ToList().JoinStrings(separator).Trim();
        }

        public static bool IsNotEmpty(this string source)
        {
            return !string.IsNullOrWhiteSpace(source);
        }

        public static bool IsEmpty(this string source)
        {
            return string.IsNullOrWhiteSpace(source);
        }

        public static string[] SplitEx(this string source, string separator = "")
        {
            return source.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
    public static class XmlNodeExtension
    {
        public static IEnumerable<XmlNode> Select(this XmlNode source,string xpath=".")
        {
            return source.SelectNodes(xpath).ToEnumerable();
        }
    }

    public static class IEnumerableExtension
    {
        public static string JoinStrings<T>(this IEnumerable<T> source,string separator="")
        {
            return string.Join(separator, source);
        }
    }
}
