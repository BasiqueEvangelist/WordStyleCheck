namespace WordStyleCheck;

public static class StringEx
{
    static int min(int a, int b, int c)
    {
        return Math.Min(Math.Min(a, b), c);
    }

    public static bool CharMatch(char a, char b)
        => a == b || !char.IsLetterOrDigit(a) || !char.IsLetterOrDigit(b);

    public static int LevenshteinDistance(string a, string b)
    {
        int[] prevRow = new int[b.Length + 1];
        int[] thisRow = new int[b.Length + 1];

        // init thisRow as 
        for (int i = 0; i < prevRow.Length; i++) prevRow[i] = i;

        for (int i = 0; i < a.Length; i++)
        {
            thisRow[0] = i + 1;
            for (int j = 0; j < b.Length; j++)
            {
                var cost = CharMatch(a[i], b[j]) ? 0 : 1;
                thisRow[j + 1] = min(thisRow[j] + 1, prevRow[j + 1] + 1, prevRow[j] + cost);
            }

            (prevRow, thisRow) = (thisRow, prevRow);
        }

        return prevRow[b.Length];
    }
}