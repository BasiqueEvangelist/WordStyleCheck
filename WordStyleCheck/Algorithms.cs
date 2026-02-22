namespace WordStyleCheck;

public class Algorithms
{
    private static int Min(int a, int b, int c) { return Math.Min(Math.Min(a,b),c);}

    private static bool? LevenshteinEdgeCases(string a, string b, int maxDistance = 1)
    {
        if (a == b) return true;
        if (string.IsNullOrEmpty(a))
        {
            if (string.IsNullOrEmpty(b)) return true;
            if (b.Length <= maxDistance) return true;
        }

        if (string.IsNullOrEmpty(b))
            if (a.Length <= maxDistance) return true;

        if (Math.Abs(a.Length - b.Length) > maxDistance) return false;

        return null;
    }

    public static bool LevenshteinNeighbors(string a, string b, int maxDistance = 1)
    {
        bool? earlyResult = LevenshteinEdgeCases(a,b,maxDistance);
        if (earlyResult.HasValue) return earlyResult.Value;

        int[] thisRow = new int[b.Length + 1];
        int[] nextRow = new int[b.Length + 1];
        
        for (int i = 0; i < thisRow.Length; i++) thisRow[i] = i;

        for (int i = 0; i < a.Length; i++)
        {
            nextRow[0] = i + 1;
            bool possibleNeighbor = false;
            for (int j = 0; j < b.Length; j++)
            {
                var cost = (a[i] == b[j]) ? 0 : 1;
                nextRow[j + 1] = Min(nextRow[j] + 1, thisRow[j + 1] + 1, thisRow[j] + cost);
                if (nextRow[j + 1] <= maxDistance) possibleNeighbor = true;
            }
            if (!possibleNeighbor) return false;
            Array.Copy(nextRow, thisRow, thisRow.Length);
        }

        if (thisRow[b.Length] <= maxDistance) return true;
        return false;
    }
}