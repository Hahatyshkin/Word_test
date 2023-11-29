using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_test
{
    public static class Class1
    {
        
    }
}
public static class Extensions
{
    public static string GetString<K, V>(this IDictionary<K, V> dict)
    {
        var items = from kvp in dict
                    select kvp.Key + ":" + kvp.Value;

        return "{" + string.Join(", ", items) + "}";
    }
}

public class Example
{
    public static void Main()
    {
        Dictionary<string, int> dict = new Dictionary<string, int>() {
            {"A", 1}, {"B", 2}, {"C", 3}, {"D", 4}, {"E", 5}
        };

        string str = dict.GetString();

        Console.WriteLine(str);        // {A:1, B:2, C:3, D:4, E:5}
    }
}
