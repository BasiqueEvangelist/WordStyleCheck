using WordStyleCheck.Profiles.Conference;
using WordStyleCheck.Profiles.Gost7_32;

namespace WordStyleCheck.Profiles;

public static class ProfileStore
{
    private static readonly Dictionary<string, IProfile> _profiles =
    new() {
        ["gost-7.32"] = new Gost7_32Profile(),
        ["conference"] = new ConferenceProfile()
    };

    public static IProfile? GetProfile(string name)
    {
        return _profiles[name];
    }
}