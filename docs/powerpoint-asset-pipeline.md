# PowerPoint Asset Pipeline

Pi-Office treats PowerPoint visuals as native Office assets where the host APIs support it. The goal is to keep slide artifacts inspectable and verifiable instead of quietly inserting placeholder glyphs.

## Supported Sources

| Source | Status | Input | Insertion path | Verification path |
| --- | --- | --- | --- | --- |
| Built-in SVG icon catalog | Available | Original Pi-Office SVG geometry | Rasterized to a PowerPoint image shape through `addImage` when supported | `verify_slide_visual` can inspect the selected image shape and Pi-Office icon alt text |
| Provided image payload | Available | PNG base64 supplied by a tool result or user workflow | Inserted as a PowerPoint image shape, or used to replace a target image placeholder | `verify_slide_visual` can return the selected image shape snapshot and metadata |
| Existing PowerPoint shape snapshot | Available | PNG snapshot exported from a selected/source shape where PowerPoint APIs support it | Copied into a destination slide or target image shape | `verify_slide_visual` can inspect the destination image shape |
| Selected PowerPoint image shape | Available | Existing selected image shape without Pi-Office source metadata | Existing presentation content; source provenance cannot be inferred from Office.js shape metadata alone | `verify_slide_visual` can confirm the selected shape is an image and return its shape metadata |
| Generated image handoff | Available through composition | OpenAI image result as PNG base64 | Use `generate_image` with insertion enabled, or insert the returned base64 through `insert_slide_element` with `assetSourceId: "generated-image-base64"` | `verify_slide_visual` can inspect the selected inserted image shape and report generated-image source metadata |
| Reusable slide component | Available | Pi-Office native component IDs: `metric-card`, `quote-callout`, `section-divider` | `insert_slide_element` with operation `add_reusable_component` creates native PowerPoint shapes with component metadata | `verify_slide_visual` can classify selected component shapes from Pi-Office alt text metadata |

## Current Rules

- `search_icons` returns the built-in SVG icon catalog plus the supported asset-source contract.
- `insert_icon` uses SVG-rasterized image shapes by default. Glyph text boxes require an explicit `allowGlyphFallback=true` or `fallback: "glyph-textbox"` option.
- `generate_image` insertions tag PowerPoint image shapes with generated-image source metadata, prompt/model metadata, and accessibility alt text where PowerPoint image insertion is available.
- Existing generated image payloads can be inserted through `insert_slide_element` / `insert_inline_picture` by passing `assetSourceId: "generated-image-base64"`.
- Reusable native components can be inserted through `insert_slide_element` / `add_reusable_component` with `componentId: "metric-card"`, `"quote-callout"`, or `"section-divider"`. Use text slots such as `metricLabel`, `metricValue`, `metricDelta`, `quote`, `attribution`, `eyebrow`, `title`, and `subtitle`.
- `verify_slide_visual` is the non-mutating check after inserting icons, images, or reusable components. Select the inserted shape before running it when asset-level verification matters.
- Full slideshow-frame screenshots are not claimed here; PowerPoint verification uses Office.js slide and shape snapshot paths.
