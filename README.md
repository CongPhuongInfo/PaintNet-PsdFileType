# PsdPlugin - VB.NET port (net9.0-windows)

VB.NET port of github.com/PsdPlugin/PsdPlugin (Paint.NET's Photoshop PSD/PSB
FileType plugin), converted piece by piece across 5 parts, targeting the
**.NET 9 SDK** - matching the upstream project's own move off .NET Framework
(current upstream targets `net7.0-windows`; this port bumps that to `net9.0`).

- `PsdFile/`             - format-parsing library (headers, layers, channels,
                            compression, image resources)
- `PhotoShopFileType/`    - the actual Paint.NET plugin integration
- `Photoshop.vbproj`      - SDK-style VB project file
- `build.bat`             - thin wrapper around `dotnet build`

## Building

1. Install the [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)
   if you don't already have it (`dotnet --version` should print `9.x`).
2. Open `Photoshop.vbproj` and check the `PaintNetDir` value (near the top)
   points at your Paint.NET install - default is `C:\Program Files\paint.net`.
   The project references `PaintDotNet.Base.dll`, `PaintDotNet.Core.dll`,
   `PaintDotNet.Data.dll`, `PaintDotNet.Primitives.dll`,
   `PaintDotNet.ComponentModel.dll`, `PaintDotNet.Resources.dll`, and
   `System.Drawing.Common.dll` from that folder.
3. Run `build.bat`, or `dotnet build Photoshop.vbproj -c Release` directly.
4. Copy the resulting `Photoshop.dll` (under
   `bin\Release\net9.0-windows\`) into Paint.NET's `FileTypes` plugin folder.

### If `RootNamespace` causes double-nested namespaces

SDK-style VB projects normally wrap every file's `Namespace` in a project-wide
root namespace. Every file here already declares its full namespace
explicitly (`PhotoshopFile`, `PaintDotNet.Data.PhotoshopFileType`) to match
the original C#, so `Photoshop.vbproj` sets `<RootNamespace></RootNamespace>`
(empty) to disable that wrapping. If your VB toolchain still prefixes
namespaces unexpectedly, check that setting first.

## Porting notes (things that differ from the original C#)

VB.NET has no `unsafe`/pointer support, so every place the original used
`fixed`/`byte*`/`ColorBgra*` was rewritten with safe array indexing instead:

- Endian swaps (`PsdBinaryReader`/`Writer`, `Util`) -> `BitConverter` +
  `Array.Reverse` on byte arrays.
- RLE encode/decode (`RleReader`/`RleWriter`) -> plain array loops.
- ZIP-with-prediction delta coding (`ZipPredict16Image`/`32Image`) ->
  `BitConverter.ToUInt16/ToInt32`/`GetBytes` at explicit byte offsets, with
  manual wraparound (`Mod`) where the original relied on unchecked integer
  overflow.
- Paint.NET `Surface.GetRowPointer` pixel access (`ImageDecoderPdn.vb`,
  `PsdSave.vb`) -> the `Surface`'s safe `(x, y)` indexer property, decoding
  each row into a scratch `ColorBgra()` array (struct array elements stay
  directly mutable, e.g. `destRow(i).R = ...`, just like the original
  `pDest->R = ...`) before writing the row back to the surface once.

Everything was verified by hand (line-by-line comparison against the C#
source); there was no VB.NET compiler available in the environment this was
written in, so please build and report back any compile errors - with code
this size, a few small mismatches (types, narrowing conversions) are normal
and quick to fix.
