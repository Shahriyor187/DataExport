#!/usr/bin/env bash
set -euo pipefail

# Run from the repository root, after packing DataExport. Do not use the
# public NuGet package as a fallback: only the exact candidate can be installed.
repo="$(pwd)"
version="$(dotnet msbuild DataExport.csproj -getProperty:Version -nologo | tr -d '\r\n')"
[[ -n "$version" ]] || { echo "Cannot determine package version"; exit 1; }
test -f "$repo/artifacts/DataExport.$version.nupkg" || {
  echo "Missing local candidate package DataExport.$version.nupkg"; exit 1;
}

tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT
cat > "$tmp/NuGet.Config" <<EOF
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="local-candidate" value="$repo/artifacts" />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
  </packageSources>
  <packageSourceMapping>
    <packageSource key="local-candidate">
      <package pattern="DataExport" />
    </packageSource>
    <packageSource key="nuget.org">
      <package pattern="*" />
    </packageSource>
  </packageSourceMapping>
</configuration>
EOF

for target in net8.0 net10.0; do
  echo "==> Testing packed DataExport $version on $target"
  folder="$tmp/$target"
  dotnet new console --framework "$target" --output "$folder" --no-restore
  dotnet add "$folder/$target.csproj" package DataExport \
    --version "$version" --no-restore
  cp tests/DataExport.PackageConsumer/Program.cs "$folder/Program.cs"
  dotnet restore "$folder/$target.csproj" --configfile "$tmp/NuGet.Config" \
    --packages "$tmp/packages" --force --no-cache
  dotnet run --project "$folder/$target.csproj" \
    --configuration Release --no-restore
done
