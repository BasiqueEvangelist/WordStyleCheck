using WordStyleCheck.Profiles.Gost7_32;
using WordStyleCheck.Profiles.IkbConference;
using WordStyleCheck.Profiles.Ntk;

namespace WordStyleCheck.Profiles;

public static class ProfileStore
{
    private static readonly Dictionary<string, IProfile> _profiles = new();

    static ProfileStore()
    {
        AddProfile(new Gost7_32Profile());
        AddProfile(new IkbConferenceProfile());
        AddProfile(new NtkProfile());
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