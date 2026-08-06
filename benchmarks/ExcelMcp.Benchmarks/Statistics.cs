namespace Sbroenne.ExcelMcp.Benchmarks;

internal sealed record ConfidenceInterval(double Low, double High);

internal sealed record DistributionSummary(
    int Count,
    double Minimum,
    double Median,
    double P95,
    double P99,
    double Maximum,
    double Mean,
    double StandardDeviation,
    ConfidenceInterval MedianConfidence95);

internal sealed record ReliabilitySummary(
    int Successes,
    int Failures,
    double SuccessRate,
    double FailureRate,
    ConfidenceInterval FailureRateConfidence95);

internal sealed record PairedComparison(
    int PairCount,
    bool LowerIsBetter,
    bool Improved,
    double BaselineMedian,
    double CandidateMedian,
    double CandidateToBaselineRatio,
    double PercentImprovement);

internal static class Statistics
{
    public static DistributionSummary Summarize(
        IReadOnlyList<double> samples,
        int bootstrapIterations = 10_000,
        int seed = 20260805)
    {
        ArgumentNullException.ThrowIfNull(samples);
        if (samples.Count == 0)
        {
            throw new ArgumentException("At least one sample is required.", nameof(samples));
        }

        ArgumentOutOfRangeException.ThrowIfLessThan(bootstrapIterations, 1);

        var sorted = samples.Order().ToArray();
        var mean = sorted.Average();
        var variance = sorted.Length > 1
            ? sorted.Sum(value => Math.Pow(value - mean, 2)) / (sorted.Length - 1)
            : 0;

        var bootstrappedMedians = BootstrapMedians(sorted, bootstrapIterations, seed);
        return new DistributionSummary(
            sorted.Length,
            sorted[0],
            Percentile(sorted, 0.50),
            Percentile(sorted, 0.95),
            Percentile(sorted, 0.99),
            sorted[^1],
            mean,
            Math.Sqrt(variance),
            new ConfidenceInterval(
                Percentile(bootstrappedMedians, 0.025),
                Percentile(bootstrappedMedians, 0.975)));
    }

    public static ReliabilitySummary SummarizeReliability(int successes, int failures)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(successes);
        ArgumentOutOfRangeException.ThrowIfNegative(failures);

        var total = successes + failures;
        if (total == 0)
        {
            throw new ArgumentException("At least one reliability observation is required.");
        }

        var failureRate = (double)failures / total;
        ConfidenceInterval interval;
        if (failures == 0)
        {
            // The rule of three is the conventional 95% upper bound when zero events are observed.
            interval = new ConfidenceInterval(0, Math.Min(1, 3d / total));
        }
        else if (successes == 0)
        {
            interval = new ConfidenceInterval(Math.Max(0, 1 - (3d / total)), 1);
        }
        else
        {
            interval = WilsonInterval(failures, total);
        }

        return new ReliabilitySummary(
            successes,
            failures,
            (double)successes / total,
            failureRate,
            interval);
    }

    public static PairedComparison ComparePaired(
        IReadOnlyList<double> baseline,
        IReadOnlyList<double> candidate,
        bool lowerIsBetter)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(candidate);
        if (baseline.Count == 0 || baseline.Count != candidate.Count)
        {
            throw new ArgumentException("Baseline and candidate must contain the same non-zero number of paired samples.");
        }

        var baselineMedian = Percentile(baseline.Order().ToArray(), 0.50);
        var candidateMedian = Percentile(candidate.Order().ToArray(), 0.50);
        if (baselineMedian == 0)
        {
            throw new ArgumentException("The baseline median must be non-zero.", nameof(baseline));
        }

        var ratio = candidateMedian / baselineMedian;
        var improvement = lowerIsBetter ? (1 - ratio) * 100 : (ratio - 1) * 100;
        return new PairedComparison(
            baseline.Count,
            lowerIsBetter,
            improvement > 0,
            baselineMedian,
            candidateMedian,
            ratio,
            improvement);
    }

    public static double Percentile(IReadOnlyList<double> sortedSamples, double percentile)
    {
        if (sortedSamples.Count == 0)
        {
            throw new ArgumentException("At least one sample is required.", nameof(sortedSamples));
        }

        if (percentile is < 0 or > 1)
        {
            throw new ArgumentOutOfRangeException(nameof(percentile));
        }

        if (sortedSamples.Count == 1)
        {
            return sortedSamples[0];
        }

        var position = (sortedSamples.Count - 1) * percentile;
        var lowerIndex = (int)Math.Floor(position);
        var upperIndex = (int)Math.Ceiling(position);
        if (lowerIndex == upperIndex)
        {
            return sortedSamples[lowerIndex];
        }

        var fraction = position - lowerIndex;
        return sortedSamples[lowerIndex] + ((sortedSamples[upperIndex] - sortedSamples[lowerIndex]) * fraction);
    }

    // Deterministic pseudorandom sampling is required for reproducible statistics;
    // this is not a security-sensitive use of randomness.
#pragma warning disable CA5394
    private static double[] BootstrapMedians(double[] source, int iterations, int seed)
    {
        var random = new Random(seed);
        var resample = new double[source.Length];
        var medians = new double[iterations];
        for (var iteration = 0; iteration < iterations; iteration++)
        {
            for (var index = 0; index < source.Length; index++)
            {
                resample[index] = source[random.Next(source.Length)];
            }

            Array.Sort(resample);
            medians[iteration] = Percentile(resample, 0.50);
        }

        Array.Sort(medians);
        return medians;
    }
#pragma warning restore CA5394

    private static ConfidenceInterval WilsonInterval(int events, int total)
    {
        const double z = 1.959963984540054;
        var proportion = (double)events / total;
        var denominator = 1 + ((z * z) / total);
        var center = (proportion + ((z * z) / (2 * total))) / denominator;
        var halfWidth = z * Math.Sqrt(
            ((proportion * (1 - proportion)) / total) + ((z * z) / (4d * total * total))) / denominator;
        return new ConfidenceInterval(
            Math.Max(0, center - halfWidth),
            Math.Min(1, center + halfWidth));
    }
}
