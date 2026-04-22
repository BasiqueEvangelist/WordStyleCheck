using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck;

public interface IProfile
{
    public List<IClassifier> Classifiers { get; }

    public List<ILint> Lints { get; }
}