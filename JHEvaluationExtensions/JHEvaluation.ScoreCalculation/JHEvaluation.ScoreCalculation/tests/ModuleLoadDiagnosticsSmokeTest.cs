using System;
using System.IO;
using System.Linq;

namespace JHEvaluation.ScoreCalculation
{
    internal enum ModuleMode
    {
        HsinChu,
        KaoHsiung
    }

    internal static class ModuleLoadDiagnosticsSmokeTest
    {
        private static int Main(string[] args)
        {
            if (args.Length != 3)
                throw new ArgumentException("Expected paths for enabled, disabled, and failure metrics.");

            VerifyEnabledMetrics(args[0]);
            VerifyDisabledDetails(args[1]);
            VerifyFailureRethrow(args[2]);
            VerifyUnwritablePathDoesNotThrow();
            Console.WriteLine("ModuleLoadDiagnostics smoke tests passed.");
            return 0;
        }

        private static void VerifyEnabledMetrics(string path)
        {
            Configure(path, "1", "optimized");
            ModuleLoadDiagnostics diagnostics = ModuleLoadDiagnostics.Start();
            diagnostics.SetDeploymentMode(ModuleMode.HsinChu);
            diagnostics.Measure("successful-stage", delegate { });
            diagnostics.Complete(true, null);
            diagnostics.Flush();

            string[] lines = File.ReadAllLines(path);
            Assert(lines.Length == 3, "Enabled diagnostics should write a header, stage, and total.");
            Assert(lines[1].Contains("\"successful-stage\"") && lines[1].Contains(",true,"), "Successful stage was not recorded correctly.");
            Assert(lines[1].Contains("\"HsinChu\""), "Deployment mode was not applied to stages measured before mode resolution.");
            Assert(lines[2].Contains("\"total\"") && lines[2].Contains(",true,"), "Successful total was not recorded correctly.");
        }

        private static void VerifyDisabledDetails(string path)
        {
            Configure(path, "0", "optimized");
            ModuleLoadDiagnostics diagnostics = ModuleLoadDiagnostics.Start();
            diagnostics.SetDeploymentMode(ModuleMode.KaoHsiung);
            diagnostics.Measure("hidden-stage", delegate { });
            diagnostics.Complete(true, null);
            diagnostics.Flush();

            string[] lines = File.ReadAllLines(path);
            Assert(lines.Length == 2, "Disabled details should write only a header and total.");
            Assert(lines[1].Contains("\"total\"") && !lines.Any(line => line.Contains("hidden-stage")), "Detailed stage was emitted while disabled.");
        }

        private static void VerifyFailureRethrow(string path)
        {
            Configure(path, "1", "optimized");
            ModuleLoadDiagnostics diagnostics = ModuleLoadDiagnostics.Start();
            Exception observed = null;
            try
            {
                diagnostics.Measure("failing-stage", delegate { throw new InvalidOperationException("expected failure"); });
            }
            catch (InvalidOperationException ex)
            {
                observed = ex;
            }

            Assert(observed != null, "Measure swallowed the initialization exception.");
            diagnostics.Complete(false, observed);
            diagnostics.Flush();

            string[] lines = File.ReadAllLines(path);
            Assert(lines.Length == 3, "Failure diagnostics should write a header, stage, and total.");
            Assert(lines[1].Contains("\"failing-stage\"") && lines[1].Contains(",false,"), "Failed stage was not recorded correctly.");
            Assert(lines[2].Contains("\"total\"") && lines[2].Contains(",false,"), "Failed total was not recorded correctly.");
            Assert(lines[1].Contains("expected failure"), "Failure details are missing.");
        }

        private static void VerifyUnwritablePathDoesNotThrow()
        {
            Configure(Path.GetPathRoot(Environment.SystemDirectory), "1", "optimized");
            ModuleLoadDiagnostics diagnostics = ModuleLoadDiagnostics.Start();
            diagnostics.Measure("successful-stage", delegate { });
            diagnostics.Complete(true, null);
            diagnostics.Flush();
        }

        private static void Configure(string path, string enabled, string variant)
        {
            if (File.Exists(path))
                File.Delete(path);

            Environment.SetEnvironmentVariable(ModuleLoadDiagnostics.MetricsPathEnvironmentVariable, path);
            Environment.SetEnvironmentVariable(ModuleLoadDiagnostics.DiagnosticsEnabledEnvironmentVariable, enabled);
            Environment.SetEnvironmentVariable(ModuleLoadDiagnostics.VariantEnvironmentVariable, variant);
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
                throw new InvalidOperationException(message);
        }
    }
}
