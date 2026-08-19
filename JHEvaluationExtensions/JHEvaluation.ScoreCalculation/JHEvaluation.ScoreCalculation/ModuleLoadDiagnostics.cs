using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace JHEvaluation.ScoreCalculation
{
    /// <summary>
    /// Collects module-load timings without introducing a remote logging dependency.
    /// Metrics are emitted after the measured load path has completed.
    /// </summary>
    internal sealed class ModuleLoadDiagnostics
    {
        internal const string DiagnosticsEnabledEnvironmentVariable = "JH_SCORE_LOAD_DIAGNOSTICS";
        internal const string MetricsPathEnvironmentVariable = "JH_SCORE_LOAD_METRICS_PATH";
        internal const string VariantEnvironmentVariable = "JH_SCORE_LOAD_VARIANT";

        private readonly List<ModuleLoadMetric> _metrics = new List<ModuleLoadMetric>();
        private readonly Stopwatch _totalStopwatch = Stopwatch.StartNew();
        private readonly bool _detailsEnabled;
        private readonly string _sessionId = Guid.NewGuid().ToString("N");
        private readonly string _variant;
        private string _deploymentMode = "Unknown";
        private bool _completed;

        private ModuleLoadDiagnostics()
        {
            _detailsEnabled = !string.Equals(
                Environment.GetEnvironmentVariable(DiagnosticsEnabledEnvironmentVariable),
                "0",
                StringComparison.OrdinalIgnoreCase);
            _variant = Environment.GetEnvironmentVariable(VariantEnvironmentVariable) ?? "unspecified";
        }

        internal static ModuleLoadDiagnostics Start()
        {
            return new ModuleLoadDiagnostics();
        }

        internal void SetDeploymentMode(ModuleMode mode)
        {
            _deploymentMode = mode.ToString();
            foreach (ModuleLoadMetric metric in _metrics)
                metric.SetDeploymentMode(_deploymentMode);
        }

        internal void Measure(string stage, Action action)
        {
            if (action == null)
                throw new ArgumentNullException("action");

            Stopwatch stopwatch = Stopwatch.StartNew();
            try
            {
                action();
                stopwatch.Stop();
                AddMetric(stage, stopwatch.Elapsed, true, null);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                AddMetric(stage, stopwatch.Elapsed, false, ex);
                throw;
            }
        }

        internal void Complete(bool success, Exception exception)
        {
            if (_completed)
                return;

            _completed = true;
            _totalStopwatch.Stop();
            AddMetric("total", _totalStopwatch.Elapsed, success, exception, true);
        }

        internal void Flush()
        {
            if (_metrics.Count == 0)
                return;

            foreach (ModuleLoadMetric metric in _metrics)
                Trace.WriteLine(metric.ToTraceMessage(), "JHEvaluation.ScoreCalculation.Load");

            string metricsPath = Environment.GetEnvironmentVariable(MetricsPathEnvironmentVariable);
            if (string.IsNullOrWhiteSpace(metricsPath))
                return;

            try
            {
                bool writeHeader = !File.Exists(metricsPath) || new FileInfo(metricsPath).Length == 0;
                List<string> lines = new List<string>();
                if (writeHeader)
                    lines.Add("TimestampUtc,SessionId,Variant,DeploymentMode,Stage,ElapsedMilliseconds,Success,Exception");

                foreach (ModuleLoadMetric metric in _metrics)
                    lines.Add(metric.ToCsv());

                File.AppendAllLines(metricsPath, lines.ToArray());
            }
            catch (Exception ex)
            {
                // Diagnostics must never change module-load success or failure semantics.
                Trace.WriteLine(
                    "Unable to write module-load metrics: " + ex.GetType().FullName + ": " + ex.Message,
                    "JHEvaluation.ScoreCalculation.Load");
            }
        }

        private void AddMetric(string stage, TimeSpan elapsed, bool success, Exception exception)
        {
            AddMetric(stage, elapsed, success, exception, false);
        }

        private void AddMetric(string stage, TimeSpan elapsed, bool success, Exception exception, bool required)
        {
            if (!_detailsEnabled && !required)
                return;

            _metrics.Add(new ModuleLoadMetric(
                DateTime.UtcNow,
                _sessionId,
                _variant,
                _deploymentMode,
                stage,
                elapsed.TotalMilliseconds,
                success,
                exception));
        }

        private sealed class ModuleLoadMetric
        {
            private readonly DateTime _timestampUtc;
            private readonly string _sessionId;
            private readonly string _variant;
            private string _deploymentMode;
            private readonly string _stage;
            private readonly double _elapsedMilliseconds;
            private readonly bool _success;
            private readonly string _exception;

            internal ModuleLoadMetric(
                DateTime timestampUtc,
                string sessionId,
                string variant,
                string deploymentMode,
                string stage,
                double elapsedMilliseconds,
                bool success,
                Exception exception)
            {
                _timestampUtc = timestampUtc;
                _sessionId = sessionId;
                _variant = variant;
                _deploymentMode = deploymentMode;
                _stage = stage;
                _elapsedMilliseconds = elapsedMilliseconds;
                _success = success;
                _exception = exception == null
                    ? string.Empty
                    : exception.GetType().FullName + ": " + exception.Message;
            }

            internal string ToTraceMessage()
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    "session={0}; variant={1}; mode={2}; stage={3}; elapsed_ms={4:F3}; success={5}; exception={6}",
                    _sessionId,
                    _variant,
                    _deploymentMode,
                    _stage,
                    _elapsedMilliseconds,
                    _success,
                    _exception);
            }

            internal void SetDeploymentMode(string deploymentMode)
            {
                _deploymentMode = deploymentMode;
            }

            internal string ToCsv()
            {
                return string.Join(",", new[]
                {
                    EscapeCsv(_timestampUtc.ToString("o", CultureInfo.InvariantCulture)),
                    EscapeCsv(_sessionId),
                    EscapeCsv(_variant),
                    EscapeCsv(_deploymentMode),
                    EscapeCsv(_stage),
                    _elapsedMilliseconds.ToString("F3", CultureInfo.InvariantCulture),
                    _success ? "true" : "false",
                    EscapeCsv(_exception)
                });
            }

            private static string EscapeCsv(string value)
            {
                string safeValue = value ?? string.Empty;
                return "\"" + safeValue.Replace("\"", "\"\"") + "\"";
            }
        }
    }
}
