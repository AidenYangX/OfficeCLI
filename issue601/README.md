# OfficeCLI #601 verification

This branch contains review artifacts only. It is not part of the proposed upstream code change.

Baseline: `iOfficeAI/OfficeCLI` commit `c56512c8` (1.0.160).
Fix: `9d1c52b9` on `AidenYangX/OfficeCLI:fix/pptx-flipped-text`.
Verified on Windows x64, .NET SDK 10.0.400, Microsoft PowerPoint 16.0, and real Chrome via OpenCLI.

The screenshots show four rectangles with unequal left/right text insets. The labels are `(flipH)/(flipV)`.

| Case | PowerPoint text orientation | Baseline previews | Fixed previews |
|---|---|---|---|
| false/false | Normal | Normal | Normal |
| true/false | Normal | Mirrored | Normal |
| false/true | Rotated 180 degrees | Mirrored vertically | Rotated 180 degrees |
| true/true | Rotated 180 degrees | Rotated 180 degrees | Rotated 180 degrees |

- [Before: Chrome HTML](before.png)
- [After: Chrome HTML](after.png)
- [Reference: PowerPoint export](powerpoint.png)

The four rectangles retain their 1 cm leading text inset within a 14 cm width. Chrome measurements verify the inset within 0.002 of the normalized width.

Additional six-shape fixture: normal, horizontal, vertical, both flips, rotation 25 degrees, and flipH plus rotation 25 degrees. HTML and SVG each pass 6/6 orientation checks after the fix, versus 3/6 before. Four asymmetric-inset cases pass in each renderer. Six HTML controls also pass: text rotation, vertical writing, and Arabic with/without flipH. Total: 26 browser orientation checks, plus inset-position checks. Native exports were visually compared.

## Reproduction

From a source checkout, build with `dotnet build src/officecli/officecli.csproj -c Release --nologo`. Then:

```powershell
python repro.py path/to/officecli.exe output-directory
```

This writes a six-shape `flips.pptx`, HTML, and SVG. Open both previews. With Microsoft PowerPoint installed, export a reference image without saving changes to the presentation:

```powershell
./export-powerpoint.ps1 -Deck output-directory/flips.pptx -Output powerpoint-reference.png
```

Every generated PPTX fixture passed `officecli validate`. Both builds had the same two existing warnings and zero errors.

The existing HTML chevron geometry approximation is outside this patch. SVG body rotation/vertical writing and group-level flips are also outside its scope.
