using System.IO.Hashing;
using DocumentFormat.OpenXml;

namespace WordStyleCheck;

public static class HashUtils
{
    public static void HashElement(OpenXmlElement element, NonCryptographicHashAlgorithm hasher)
    {
        if (element.Parent == null)
        {
            hasher.Append([0]);
            return;
        }

        int index = element.Parent.ChildElements.ToList().IndexOf(element);
        hasher.Append(BitConverter.GetBytes(index));
        HashElement(element.Parent, hasher);
    }
}