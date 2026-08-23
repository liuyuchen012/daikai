#!/bin/bash
set -e
echo "=== Installing .NET SDK 10.0 ==="
curl -sL https://dot.net/v1/dotnet-install.sh -o /tmp/dotnet-install.sh
chmod +x /tmp/dotnet-install.sh
bash /tmp/dotnet-install.sh --channel 10.0 --install-dir $HOME/.dotnet
export PATH="$HOME/.dotnet:$PATH"
echo "=== .NET version ==="
dotnet --version
echo "=== Publishing Server ==="
PROJ="./check-in-net/Server/CheckIn.Server.csproj"
OUT="./check-in-net/release/Server-linux"
dotnet publish "$PROJ" -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -o "$OUT"
echo "=== Done ==="
ls -la "$OUT/CheckIn.Server"
