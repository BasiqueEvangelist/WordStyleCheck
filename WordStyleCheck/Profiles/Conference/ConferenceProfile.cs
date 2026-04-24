using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles.Conference;

public class ConferenceProfile : IProfile
{
    public List<IClassifier> Classifiers { get; } =
    [
    ];

    public List<ILint> Lints { get; } =
    [
        // new FontSizeLint(x => x is {}, 28, false, "IncorrectFontSize"),
    ];
}