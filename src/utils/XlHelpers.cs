using System.Xml;

namespace SPSReader.utils;

public static class XlHelpers
{
    public static async Task<bool> ReadToFollowingAsync(XmlReader reader, string localName)
    {
        while (await reader.ReadAsync())
        {
            if (reader.NodeType == XmlNodeType.Element && 
                (reader.LocalName == localName || reader.Name == localName))
            {
                return true;
            }
        }
        return false;
    }

    public static async Task<bool> ReadToDescendantAsync(XmlReader reader, string localName)
    {
        if (reader.IsEmptyElement) return false;

        var depth = reader.Depth;
        while (await reader.ReadAsync())
        {
            if (reader.Depth <= depth)
                return false;

            if (reader.NodeType == XmlNodeType.Element && 
                (reader.LocalName == localName || reader.Name == localName))
            {
                return true;
            }
        }

        return false;
    }
}