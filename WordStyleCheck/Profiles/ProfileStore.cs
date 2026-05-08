using WordStyleCheck.Profiles.Conference;
using WordStyleCheck.Profiles.Gost7_32;

namespace WordStyleCheck.Profiles;

public static class ProfileStore
{
    private static readonly Dictionary<string, IProfile> _profiles = new();

    static ProfileStore()
    {
        AddProfile(new Gost7_32Profile());
        AddProfile(new ConferenceProfile());
    }
    
    public static readonly IEnumerable<IProfile> Profiles = _profiles.Values;
    
    public static IProfile? GetProfile(string name)
    {
        return _profiles[name];
    }

    public static void AddProfile(IProfile profile)
    {
        _profiles.Add(profile.Id, profile);
    }
}