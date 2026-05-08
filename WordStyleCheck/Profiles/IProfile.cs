using WordStyleCheck.Analysis;
using WordStyleCheck.Lints;

namespace WordStyleCheck.Profiles;

public interface IProfile
{
    public string Id { get; }
    
    public string Name { get; }
    
    public List<IClassifier> Classifiers { get; }

    public List<ILint> Lints { get; }
}