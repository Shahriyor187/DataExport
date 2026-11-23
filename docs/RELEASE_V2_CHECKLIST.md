# DataExport v2 release readiness

This document tracks the proposed **2.0.0** major release. A GitHub
`2.0.0-rc.1` source candidate is not a published NuGet package.
Do not publish without explicit user approval.

## Shipping gates

- [x] Verify .NET 8 and .NET 10 test suites on the RC commit:
      28 tests passed on each target in Actions run 45191053994.
      Repeat these checks on the future stable release commit.
- [x] Streaming XLSX regression suite and 10k-row CI smoke validation passed:
      workflow run 45191053943. Revalidate on the stable release commit.
- [x] Built DataExport.2.0.0-rc.1.nupkg and .snupkg. CI inspected
      README, license, package ID and net8.0/net10.0 binaries
      (run 45191053994). Verify the exact stable package separately.
- [x] Installed the exact local RC package into isolated .NET 8 and
      .NET 10 console apps; executed JSON, XML, CSV, buffered XLSX,
      direct-ZIP streaming XLSX and multi-sheet report checks successfully
      (run 45191053994). Repeat on the stable package.
- [x] Automated audit completed: 16 resolved package entries,
      no reported direct/transitive advisories (run 45191053994).
      A point-in-time scan is not a comprehensive security review.
- [x] Reviewed compatibility intent: the current `DataExporter` is sealed;
      System.Text.Json replaces Newtonsoft.Json; the original public
      file-based methods remain. A major 2.0.0 version communicates the
      behavioral/subclassing compatibility impact.
- [ ] Verify large-data resource usage with representative datasets; recorded
      100k and 1M performance checks use synthetic data on GitHub runners.
- [x] Documented stream ownership, cancellation cleanup and supported
      XLSX features. The direct-ZIP writer does not implement Excel charts,
      calculated formulas or a full in-memory worksheet feature set.
- [ ] Ensure consuming services enforce access controls, authorization,
      output quotas, request timeouts, concurrency limits, disk space and
      cleanup; this library itself cannot impose upstream hosting controls.
- [ ] Convert the source candidate version to stable `2.0.0`, re-run
      every gate on the exact stable commit, and update changelog/README.
- [ ] Obtain explicit approval before triggering guarded manual NuGet publish.
- [ ] Verify NuGet.org listing and create GitHub tag/release only after
      a successful publish.

## CI evidence for candidate (2026-11-23)

- [DataExport CI, RC SHA b501d66b](https://github.com/Shahriyor187/DataExport/actions/runs/45191053994):
  **28/28 tests passed on net8.0**, **28/28 on net10.0**,
  exact local package `DataExport.2.0.0-rc.1.nupkg` verified and
  executable consumers passed on both target frameworks.
  16 resolved dependencies were audited with no reported vulnerabilities.
- [Streaming XLSX performance check](https://github.com/Shahriyor187/DataExport/actions/runs/45191053943):
  10,000-row smoke workload completed successfully.

These runs apply to the **source release candidate only**, not the future
stable NuGet publication. A later version change creates a new commit
requiring fresh CI and package checks.

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
