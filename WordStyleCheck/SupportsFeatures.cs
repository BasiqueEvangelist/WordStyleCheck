namespace WordStyleCheck;

public abstract class SupportsFeatures<TSelf> where TSelf : SupportsFeatures<TSelf>
{
    private readonly Dictionary<IFeatureKey, object> _features = [];

    public T? GetFeature<T>(FeatureKey<T, TSelf> key) => (T?)_features.GetValueOrDefault(key);

    public void SetFeature<T>(FeatureKey<T, TSelf> key, T value) => _features[key] = value ?? throw new ArgumentNullException(nameof(value));
}

interface IFeatureKey;

public class FeatureKey<T, THost> : IFeatureKey where THost : SupportsFeatures<THost>;