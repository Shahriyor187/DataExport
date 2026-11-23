# DataExport v2 release readiness

This document tracks the proposed **2.0.0** major release. A GitHub
`2.0.0-rc.1` source candidate is not a published NuGet package.
Do not publish without explicit user approval.

## Shipping gates

- [ ] Verify .NET 8 and .NET 10 test suites on the exact release commit.
- [ ] Verify the streaming XLSX regression suite and 10k-row smoke test.
- [ ] Build `.nupkg` and `.snupkg`, inspect license, README, package version,
      and both target-framework binaries.
- [ ] Install **only the packed local candidate** into clean .NET 8 and
      .NET 10 console apps. Execute JSON, XML, CSV, buffered XLSX,
      direct-ZIP streaming XLSX and multi-sheet report outputs.
- [ ] Run and review direct and transitive NuGet advisory checks.
- [ ] Confirm API compatibility expectations and major version: the
      `DataExporter` class is now sealed; System.Text.Json replaces
      Newtonsoft.Json; existing methods `ToJson/ToXml/ToExcel/ToCsv` remain.
- [ ] Verify large-data resource usage with representative datasets; recorded
      100k and 1M performance checks use synthetic data on GitHub runners.
- [ ] Review streaming limitations: invalid output must be discarded after
      cancellation/errors; output stream is owned by caller; ZIP output is
      forward-only and no chart/formula rendering is supported.
- [ ] Ensure consuming services enforce access controls, authorization,
      output quotas, request timeouts, concurrency limits, disk space and
      cleanup; this library itself cannot impose upstream hosting controls.
- [ ] Convert the source candidate version to stable `2.0.0`, re-run
      every gate on the exact stable commit, and update changelog/README.
- [ ] Obtain explicit approval before triggering guarded manual NuGet publish.
- [ ] Verify NuGet.org listing and create GitHub tag/release only after
      a successful publish.

## Packaging and safety

- The publishing workflow must be run manually from `master` in
  `Shahriyor187/DataExport` and requires the confirmation text `PUBLISH`.
  It checks version, tests, dependency advisories, package contents and
  isolated package consumers before using NuGet Trusted Publishing.
- The direct-ZIP XLSX path uses stream writes and does not hold all records
  in memory. That does not make memory use constant under every input:
  row properties, worksheet count, source reader and consumer infrastructure
  also contribute to resources.
- Microsoft .NET 8 runtime support ended 2026-11-10. `net8.0` packaging
  remains for backward compatibility, but deployment should use .NET 10
  when possible.
- The historical initial package description advertised Word/PDF, but
  the original source did not implement those export methods. Supported
  DataExport formats are JSON, XML, CSV and XLSX.
- Version change to 2.0.0 is intentional given public class sealing and
  JSON serialization semantics. Applications relying on undocumented
  output formatting should validate before upgrading.

## Performance baseline already recorded

From the 2026-11-22 synthetic GitHub .NET 10 direct-ZIP performance checks:

| Rows | Elapsed | Process peak | File size |
| ---: | ---: | ---: | ---: |
| 100,000 | 1.527 s | 65.4 MB | 4.63 MB |
| 1,000,000 | 3.860 s | 134.09 MB | 46.47 MB |

These results were generated under specific runner conditions and are not
predictive of production performance.
