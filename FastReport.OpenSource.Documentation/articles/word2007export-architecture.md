# Word2007Export Reconstruction Spec

This document is a reconstruction-oriented technical spec for the library exporter:

- `FastReport.Export.OoXML.Word2007Export`

It is based only on the library code that has already been decompiled and reviewed, not on the custom exporter in `Tools/FastReport.CustomExport`.

The goal of this document is not just to describe the exporter. The goal is to make it possible to recreate it with high confidence.

This document therefore separates the content into:

- directly confirmed behavior
- inferred behavior with strong evidence
- missing pieces that still block an exact byte-for-byte recreation

## Source Basis

The current spec is based on the decompiled contents of:

- `Word2007Export`
- `ExportMatrix`
- `ExportIEMObject`
- `ExportIEMStyle`
- `BinaryTree`
- `OoXMLBase`
- `OoXMLHeader`
- `OoXMLFooter`
- `OoXMLDocument.ExportSinglePage(...)` body fragment

## Reconstruction Goal

If you want to recreate the official exporter, the intended implementation strategy is:

1. recreate the same object model
2. recreate the same layout normalization process
3. recreate the same matrix splitting logic
4. recreate the same table-writing logic
5. only then recreate packaging and side parts

Trying to jump directly from `PreparedPage -> DOCX XML` without the matrix layer is the wrong abstraction.

## Confidence Levels

### Directly Confirmed

The following are directly confirmed from decompiled code:

- `Word2007Export` is matrix-based in normal mode
- `ParagraphBased` still depends on matrix mode
- `ExportMatrix` is the actual layout normalization engine
- page content is converted into `ExportIEMObject` + `ExportIEMStyle`
- coordinates are discretized with `BinaryTree`
- a 2D matrix is rasterized and analyzed
- objects can be split into fragments
- borders are cloned and optimized after splitting
- `OoXMLDocument.ExportSinglePage(...)` writes a `w:tbl` in matrix mode
- page size and margins are generated from matrix page metadata
- header/footer use dedicated matrices and dedicated XML parts
- watermarks are injected through header XML

### Strongly Inferred

The following are not fully visible, but the architecture is clear enough that a very close recreation is possible:

- the table writer uses matrix cell boundaries as Word table row/column boundaries
- cell merging is based on the final `ExportIEMObject.x/y/dx/dy`
- text paragraph output uses `ExportIEMStyle` fields such as padding, line height, paragraph offset, tabs, RTL, and alignment
- image export uses object geometry and writer-level media relations

### Still Missing

These pieces prevent an exact line-by-line recreation:

- full `OOExportBase`
- full `OoXMLDocument`
- `ExportIEMPage`
- `BinaryTreeNode`
- full cell merge writer code
- full paragraph/tab writer code
- full image writer code

Without these, the exporter can still be recreated very closely, but not guaranteed identical.

## Top-Level Export Architecture

The exporter has three layers:

1. orchestration layer
2. matrix normalization layer
3. OOXML writing layer

### 1. Orchestration Layer

Owned by:

- `Word2007Export`

Responsibilities:

- read export options
- create body/header/footer matrices
- feed report bands/objects into the matrices
- create OOXML parts
- write relationships and package content

### 2. Matrix Normalization Layer

Owned by:

- `ExportMatrix`
- `ExportIEMObject`
- `ExportIEMStyle`
- `BinaryTree`

Responsibilities:

- flatten the report page into export objects
- normalize coordinates
- generate unique X/Y boundaries
- assign object occupancy into a cell grid
- split objects into smaller rectangles when necessary
- preserve visual semantics like fill, border, text style, hyperlink, bookmark

### 3. OOXML Writing Layer

Owned by:

- `OoXMLDocument`
- `OoXMLHeader`
- `OoXMLFooter`
- other OOXML part classes

Responsibilities:

- convert the normalized matrix into Word XML
- write `document.xml`
- write headers and footers
- write settings/styles/font tables/content types/relationships

## Object Model

### `Word2007Export`

Important fields:

- `coreDocProp`
- `applicationProp`
- `document`
- `fontTable`
- `wordStyles`
- `footNotes`
- `endNotes`
- `headers`
- `footers`
- `wordSettings`
- `matrixList`
- `matrixBased`
- `paragraphBased`
- `disableSectionBreak`
- `wysiwyg`
- `saveRowHeight`
- `rowHeightIs`
- `rowHeight`
- `prepared_page`
- `doNotExpandShiftReturn`
- `matrix`
- `headerMatrices`
- `footerMatrices`
- `pageNo`
- `printOptimized`
- `memoryOptimized`
- `bookmarks`
- `useHeaderAndFooter`
- `headerHeight`

Reconstruction meaning:

- the class is stateful across the whole export
- one export instance accumulates per-page matrices
- the OOXML document object is also long-lived for the whole export

### `ExportMatrix`

Important fields:

- `iemObjectList`
- `iemStyleList`
- `xPos`
- `yPos`
- `pages`
- `matrix`
- `width`
- `height`
- `maxWidth`
- `maxHeight`
- `minLeft`
- `minTop`
- `deltaY`
- `inaccuracy`
- export options like `fillArea`, `images`, `wrap`, `bookmarks`, `keepRichText`, `printOptimized`

Reconstruction meaning:

- this is both a collector and a layout processor
- one matrix represents one logical export surface
- for body export this is a page body
- for header/footer export this is a header/footer surface

### `ExportIEMObject`

Fields that matter most for recreation:

- logical content:
  - `Text`
  - `OriginalText`
  - `Value`
  - `TextRenderType`
  - `ParagraphFormat`
  - `TabPositions`
  - `TabWidth`
- visual content:
  - `PictureStream`
  - `Metafile`
  - `ImageData`
  - `PicBounds`
- geometry:
  - `Left`
  - `Top`
  - `Width`
  - `Height`
- matrix geometry:
  - `x`
  - `y`
  - `dx`
  - `dy`
- style:
  - `StyleIndex`
  - `Style`
- semantic flags:
  - `IsText`
  - `IsRichText`
  - `IsBand`
  - `IsSvg`
  - `IsNumeric`
  - `IsDateTime`
  - `IsPercent`
  - `AutoSize`
  - `Exist`
- linkage:
  - `Parent`
  - `Hyperlink`
  - `Bookmark`

Reconstruction meaning:

- this is not just a rectangle
- it is the normalized export cell candidate
- after splitting, several child `ExportIEMObject` instances may share a parent

### `ExportIEMStyle`

Fields:

- `Font`
- `VAlign`
- `HAlign`
- `TextFill`
- `Fill`
- `Format`
- `Border`
- `Padding`
- `FirstTabOffset`
- `Underlines`
- `Angle`
- `RTL`
- `WordWrap`
- `LineHeight`
- `ParagraphOffset`
- `FontWidthRatio`
- `ForceJustify`
- `AutoWidth`
- `AutoShrink`
- `IsAutoBorder`

Reconstruction meaning:

- styles are deduplicated
- style equality is important for reuse
- every writer downstream should rely on style objects, not raw report objects

### `BinaryTree`

Fields:

- `root`
- `nodes`
- `count`
- `index`
- `inaccuracy`
- `maxDistance`
- `prevValue`

Responsibilities:

- insert coordinates with tolerance
- optionally inject intermediate coordinates when gaps are too large
- produce an ordered indexed edge array

Reconstruction meaning:

- this is not just a container
- it is one of the main geometry normalization algorithms

## Export Control Flow

This section describes the actual execution order of the exporter as confirmed by `Word2007Export`.

### Phase 1: `Start()`

`Word2007Export.Start()` performs global export initialization.

Confirmed steps:

1. call `base.Start()`
2. create `ZipArchive`
3. create OOXML parts:
   - core properties
   - application properties
   - main document part
   - font table
   - styles
   - settings
4. add main document relationships:
   - settings
   - styles
   - font table
5. initialize `matrixList`
6. initialize `pageNo = 0`
7. detect watermark presence from the first prepared page
8. if header/footer or paragraph mode is needed:
   - create header/footer dictionaries
   - optionally create footnotes/endnotes
   - scan all prepared pages for `PageHeaderBand` / `PageFooterBand`
   - allocate one OOXML header/footer part per unique band name
   - allocate matching relation ids in the main document
9. initialize `headerMatrices` and possibly `footerMatrices`

Important reconstruction detail:

- header/footer parts are deduplicated by band name, not blindly by page index
- body matrices are page-oriented
- header/footer matrices are keyed by page index but backed by reusable header/footer part names

### Phase 2: `ExportPageBegin(page)`

For matrix-based export this starts one page-level layout capture.

Confirmed steps:

1. reset `headerHeight = 0`
2. allocate a fresh body `ExportMatrix`
3. configure matrix options:
   - `Inaccuracy`
   - image resolution if `PrintOptimized`
   - `RowHeightIs`
   - `SaveRowHeight`
   - `PlainRich`
   - `AreaFill`
   - `CropAreaFill`
   - `Report`
   - `Images`
   - `WrapText`
   - `Bookmarks`
   - `FullTrust`
   - `KeepRichText`
   - `PrintOptimized`
4. call `matrix.AddPageBegin(page)`
5. if paragraph mode or watermark/header/footer mode is active:
   - allocate page-specific header and footer matrices
   - configure them with `ConfigureMatrix(...)`
   - call `AddPageBegin(page)` on header/footer matrices too

Observed option values used by `Word2007Export`:

- `PlainRich = true`
- `AreaFill = true`
- `CropAreaFill = false`
- `Images = true`
- `WrapText = false`
- `FullTrust = false`

These defaults matter because they define the exported semantic model.

### Phase 3: `ExportBand(band)`

This is the feed phase. Every band is routed either to the body matrix or to header/footer matrices.

Confirmed behavior:

- if matrix mode and the band is `PageHeaderBand` or `PageFooterBand` and header/footer processing is enabled:
  - call `CheckingBands(band)`
- otherwise:
  - if `UseHeaderAndFooter`, temporarily subtract `headerHeight` from `band.Top`
  - add the band to the body matrix with `matrix.AddBand(band, this)`
  - restore `band.Top` if modified

`CheckingBands(band)` does:

- for `PageHeaderBand`:
  - set `headerHeight = band.Height`
  - set `band.Top = 0`
  - add to `headerMatrices[pageNo]`
- for `PageFooterBand`:
  - set `band.Top = 0`
  - add to `footerMatrices[pageNo]`

Why this matters:

- when Word header/footer export is enabled, page header/footer layout is normalized into separate matrix coordinate systems
- the body matrix is shifted upward by header height so the body content starts at the right logical origin

### Phase 4: `ExportPageEnd(page)`

Confirmed steps:

1. detect whether this is the last logical exported page
2. if header/footer matrices exist:
   - call `AddPageEnd(page)` on them
3. call `matrix.AddPageEnd(page)` on the body matrix
4. call `matrix.Prepare()`
5. if memory-optimized:
   - stream matrix immediately through `document.ExportMatrix(...)`
   - clear matrix
   - keep it in `matrixList` for lifecycle consistency
6. otherwise:
   - store the prepared matrix in `matrixList`

The key point:

- all object collection happens before `Prepare()`
- all geometry normalization, occupancy rasterization, splitting, and border optimization happen inside `Prepare()`

### Phase 5: `Finish()`

Confirmed steps:

1. call `ExportOOXML(Stream)`
2. save the zip package to the destination stream
3. clear the zip object
4. dispose body matrices
5. dispose header/footer matrices when allocated

`ExportOOXML(Stream)` itself:

1. write headers/footers/notes if needed
2. write package-level relationships
3. write `[Content_Types].xml`
4. export app/core/style/settings/font parts
5. export main document using:
   - `document.Export(this, FileStream)` in memory-optimized matrix mode
   - `document.Export(this, matrixList)` in normal matrix mode
   - `document.ExportLayers(this)` in non-matrix mode

## ExportMatrix Collection Phase

The matrix does not start from grid cells. It starts from collected logical objects.

Collection order:

1. page metadata via `AddPageBegin(page)`
2. band rectangles via `AddBandObject(band)` when applicable
3. child objects via `band.ForEachAllConvectedObjects(sender)`
4. page finalization via `AddPageEnd(page)`
5. normalization via `Prepare()`

### `AddPageBegin(page)`

Confirmed captured page properties:

- `Landscape`
- `Width`
- `Height`
- `RawPaperSize`
- `LeftMargin`
- `TopMargin`
- `RightMargin`
- `BottomMargin`

The values stored in the matrix page are already scaled by `zoom`.

### `AddBand(band, sender, isHeader)`

This is the main flattening function.

Confirmed behavior:

- optionally skip bands depending on:
  - `dataOnly`
  - `seamless`
  - page header/footer conditions
- add the band rectangle itself if it has fill/border and is not texture-filled
- if the band fill is texture and images are allowed, export band as image
- iterate all convected objects
- normalize special types:
  - `CellularTextObject` -> owning table
  - `TableCell` is skipped as direct iteration result because tables are exported through `AddTableObject`
- route object by type:
  - memo-like text -> `AddTextObject`
  - nested band -> `AddBandObject`
  - simple line/transparent rectangle -> `AddLineObject`
  - HTML in html mode -> `AddHtml`
  - rich text when `keepRichText` -> `AddRich`
  - SVG in html mode -> `AddSVG`
  - other visual objects -> picture export

This dispatch tree is fundamental to recreation. If you do not match it, your matrix will diverge from theirs very early.

## Coordinate Capture And Geometry Normalization

The core geometry function is:

- `AddSetObjectPos(ReportComponentBase Obj, ref ExportIEMObject FObj)`

This is where raw report object geometry is converted into export-space geometry.

### Coordinate formulas

Confirmed formulas:

- `FObj.Left`
  - if `Obj.AbsLeft < 0`, use `0`
  - otherwise:
    - if width is positive: `AbsLeft * zoom`
    - if width is negative: `(AbsLeft + Width) * zoom`
- `FObj.Top`
  - if `Obj.AbsTop < 0`, use `deltaY`
  - otherwise:
    - if height is positive: `deltaY + AbsTop * zoom`
    - if height is negative: `deltaY + (AbsTop + Height) * zoom`
- `FObj.Width`
  - for normal objects: `abs(Width) * zoom`
  - for lines with zero width/height: at least `max(Inaccuracy + 0.1, Border.Width * zoom)`
- `FObj.Height`
  - same rule as width

### Right edge clipping

Confirmed behavior:

- if `Left + Width` exceeds printable width
- then width is reduced to:
  - `(page.Width - page.LeftMargin - page.RightMargin) * Units.Millimeters - Left`

This matters because matrix objects do not extend beyond the printable page width even if source geometry does.

### Global extents updated during capture

Each object updates:

- `maxWidth`
- `maxHeight`
- `minLeft`
- `minTop`
- `left`

Each object also contributes coordinate boundaries:

- `xPos.Add(Left)`
- `xPos.Add(Left + Width)`
- `yPos.Add(Top)`
- `yPos.Add(Top + Height)`

This is the seed for later grid construction.

## Style Extraction

Style extraction happens in:

- `AddStyle(ReportComponentBase Obj)`

The exporter does not carry report objects forward into OOXML writing. It first converts them into deduplicated export styles.

### Text styles

For `TextObject`, confirmed fields copied into `ExportIEMStyle`:

- font, scaled by `zoom` if needed
- text fill
- RTL
- word wrap
- fill
- format
- vertical alignment
- horizontal alignment, with RTL-aware swapping
- padding, scaled by `zoom`
- first tab offset
- border, zoomed when needed
- angle
- line height
- paragraph offset
- font width ratio
- force justify
- auto shrink
- auto width

This is the main source of text spacing behavior.

### Band styles

For bands:

- fill
- border
- `AutoShrink = None`
- `AutoWidth = false`

### Line styles

For line-like objects:

- force one border side depending on line orientation
- use the default 1pt font only as a placeholder

### Shape and generic visual styles

Depending on object type:

- border
- fill
- sometimes centered alignment defaults
- sometimes bitmap fallback later in the pipeline

### Style deduplication

`AddStyleInternal(...)` scans existing styles backward and reuses a matching one if `ExportIEMStyle.Equals(...)` returns true.

This means:

- style identity is value-based
- object count and style count are intentionally decoupled
- a reimplementation must preserve this semantics if it wants comparable output density

## Text Classification And Content Semantics

Text export is not just plain strings.

### `AddTextObject(TextObject Obj, bool isHeader)`

Confirmed captured content:

- `Text`
- `AutoSize`
- `TabPositions`
- `OriginalText` when exporting header clones
- `TabWidth`
- `ParagraphFormat` for `HtmlParagraph`
- inline image cache for HTML paragraph text
- semantic parsing:
  - `IsDateTime`
  - `IsPercent`
  - `IsNumeric`
  - parsed `Value`
- `TextRenderType`
- `Hyperlink`
- `Bookmark`

The matrix therefore carries both:

- presentation text
- parsed semantic value

That dual representation is important for exporters that preserve numeric/date semantics.

### Memo vs picture decision

`IsMemo(Obj)` returns true only when:

- object is `TextObject`
- angle is zero or rotated text is not forced to image
- text outline is disabled

Otherwise the object is treated as image content.

That means the official exporter intentionally degrades some text to image when it cannot preserve layout semantics safely.

## Table Object Expansion

Tables are flattened explicitly through:

- `AddTableObject(TableBase table)`

Confirmed behavior:

1. call `table.EmulateOuterBorder()`
2. call `table.EmulateFill()`
3. iterate rows and columns
4. skip cells inside a span
5. assign synthetic `Left` / `Top` to each cell from accumulated row/column sizes
6. export each visible table cell either as text or image

Important implication:

- table geometry becomes ordinary export object geometry before matrix rasterization
- there is no special table structure preserved at matrix level
- after flattening, table cells compete for space like any other object rectangles

## Matrix Preparation

`ExportMatrix.Prepare()` is the decisive transformation stage.

Confirmed body:

1. optionally insert a synthetic transparent area object when `fillArea` is enabled
2. call `Render()`
3. call `Analyze()`
4. optionally call `OptimizeFrames()`

That sequence is not interchangeable. Recreating the exporter means reproducing it in that order.

### Step 1: Synthetic base fill object

When `fillArea` is enabled:

- create a transparent fill style
- create an `ExportIEMObject`
- insert it at index `0` in `iemObjectList`
- set:
  - `Left = 0` / `Top = 0`, or cropped values if `cropFillArea`
  - `Width = MaxWidth`
  - `Height = MaxHeight`
  - `IsText = true`
  - `x = y = 0`
  - `dx = dy = 1`
- also insert `0` into both `xPos` and `yPos`

Purpose:

- guarantee a base occupancy layer
- ensure empty areas still belong to a defined background object
- simplify later fill inheritance logic

### Step 2: `Render()`

`Render()` converts collected export objects into:

- a closed X axis
- a closed Y axis
- a 2D occupancy array of object indices

#### 2.1 Close the coordinate trees

Confirmed:

- `xPos.Close()`
- `yPos.Close()`
- `width = xPos.Count`
- `height = yPos.Count`
- allocate `matrix = new int[width * height]`
- initialize all cells to `-1`

`BinaryTree.Close()`:

- repeatedly rebuilds the ordered node array until no new nodes are injected during indexing
- this happens because `MaxDistance` can force extra intermediate coordinates during `Index(...)`

This behavior is critical. It means final grid lines are not just original object boundaries. They may also include synthetic boundaries inserted to limit maximum cell size.

#### 2.2 Snap every object to indexed boundaries

For each `ExportIEMObject`:

1. find `x = xPos.IndexOf(Left)`
2. find `x2 = xPos.IndexOf(Left + Width)`
3. set `dx = x2 - x`
4. snap `Left` to `xPos.Nodes[x].value`
5. optionally resnap `Width`
6. do the same for Y

The optional resnap occurs when `Inaccuracy > 1`.

Implication:

- geometry is discretized before occupancy fill
- final cell spans come from snapped grid edges, not raw floating-point values

#### 2.3 Background fill inheritance

Confirmed special case:

- if object fill color is transparent
- inspect the current cell occupant at the object origin
- if there is already an underlying object with a style
- and that fill color differs from both the current fill and text color
- clone the style and replace the transparent fill with the underlying fill

This is one of the subtle visual-preservation mechanisms.

Meaning:

- transparent objects can inherit visual background from the layer below
- this reduces visual mismatches after cell segmentation

#### 2.4 Border borrowing from the bottom object

When the current cell origin already contains an object:

- compare edges with `CheckFrameFromBottmObject(...)`
- possibly copy bottom/left/right/top border segments into the current object style

The function adds a border side only when:

- the corresponding positions coincide exactly in the grid
- the current object does not already have that border side
- the lower object does have it

This is a border continuity repair pass performed before matrix fill.

#### 2.5 Fill the occupancy matrix

If the object is not a visually empty white/transparent band with no border:

- call `FillArea(x, y, dx, dy, objectIndex)`

That writes the object index into every cell it covers.

This is the rasterization step.

### Step 3: `Analyze()`

`Analyze()` walks the cell matrix and ensures each visible rectangle is represented by a concrete export object fragment.

Confirmed behavior:

1. iterate all matrix cells except last row/column edges
2. for each occupied cell:
   - get `objIndex`
   - if that object is not yet marked `Exist`
   - call `FindRect(x, y, out dx, out dy)`
   - compare discovered rectangle with object's stored `x/y/dx/dy`
   - if different:
     - call `Cut(objIndex, x, y, dx, dy)`
   - otherwise mark object as existing

### `FindRect(...)`

Confirmed algorithm:

- start at a cell
- extend horizontally while cell indices are identical
- use the first row width as initial width
- continue downward while each next row maintains at least that width
- stop at the first row that breaks the rectangle

This yields the maximal homogeneous rectangle for that object index at that position.

### `Cut(...)`

Confirmed behavior:

- create a new `ExportIEMObject`
- set `Parent = originalObject`
- copy:
  - style index / style
  - left/top/width/height from grid boundaries
  - text/image/base/hash metadata
  - paragraph metadata
  - semantic flags
- if the resulting fragment is still large enough:
  - move text/value/original text to the child fragment
  - clear them on the original object
- add new object to the object list
- replace matching rectangle cells from old object index to new object index
- clone borders appropriately with `CloneFrames(...)`
- mark child fragment as `Exist`

The important design rule:

- object splitting is destructive with respect to content ownership
- only one fragment keeps the text payload
- other fragments become visual fragments that preserve fill/border occupancy

This is likely one of the core reasons the exporter can represent complex overlapping rectangles as table cells.

### Step 4: `OptimizeFrames()`

This pass removes duplicated borders between adjacent cells.

Confirmed logic:

- for each occupied cell
- if the object has borders:
  - compare top border with bottom border of the cell above
  - compare left border with right border of the cell to the left
  - if the adjacent border has same width and color
  - remove the redundant border from the current object by cloning style and dropping that side

This is a table de-duplication optimization. Without it, adjacent cells would draw double lines.

## Exact Spacing Systems

The exporter has multiple spacing systems. Recreating the code means not collapsing them into one.

### 1. Geometric spacing between objects

This is the largest spacing system.

It comes from:

- `AbsLeft`
- `AbsTop`
- `Width`
- `Height`
- `deltaY`
- clipping to printable width

This spacing determines:

- where grid lines are created
- how many rows and columns the matrix has
- the raw empty areas between objects

### 2. Edge snapping tolerance

This comes from:

- `BinaryTree.Inaccuracy`

Behavior:

- two nearby coordinates within tolerance collapse to one grid line
- larger tolerance reduces cell count and increases editability
- smaller tolerance preserves visual fidelity

`Word2007Export` sets matrix inaccuracy differently depending on mode:

- not WYSIWYG: `10`
- paragraph-based: `5`
- normal WYSIWYG matrix mode: `0.3`

This is one of the exporter's main fidelity knobs.

### 3. Maximum cell size insertion

This comes from:

- `BinaryTree.MaxDistance`

Behavior:

- when consecutive coordinates are too far apart
- `BinaryTree.Index(...)` inserts intermediate synthetic nodes

Effect:

- large empty spans can be subdivided
- writer gets bounded row heights / column widths
- downstream table layout becomes more stable

We do not yet see where `MaxDistance` is set for Word export, but the tree supports it and the reconstruction should preserve that mechanism.

### 4. Row height policy

This comes from:

- `SaveRowHeight`
- `RowHeightIs`
- `RowHeight`

Directly confirmed:

- these values are pushed from `Word2007Export` into `ExportMatrix`

Strongly inferred:

- `OoXMLDocument.Export_TableRow(...)` uses them to emit:
  - exact height
  - minimum height
  - or omit explicit row height depending on settings

This is one of the currently missing exact writer details, but the intent is unambiguous.

### 5. Text internal spacing

Text spacing comes from `ExportIEMStyle`:

- `Padding`
- `LineHeight`
- `ParagraphOffset`
- `FirstTabOffset`
- `TabPositions`
- `TabWidth`
- `WordWrap`
- `HAlign`
- `VAlign`
- `RTL`
- `FontWidthRatio`
- `ForceJustify`
- `AutoShrink`
- `AutoWidth`

This is not grid spacing. This is spacing inside the exported cell content.

### 6. Page spacing

Page-level section spacing comes from matrix page metadata converted in `OoXMLDocument.ExportSinglePage(...)`.

Confirmed formulas:

- width in twips-like Word units:
  - `(PageWidth * 567 + 4) / 10`
- height:
  - `(PageHeight * 567 + 4) / 10`
- top margin:
  - `(PageTMargin * 530 + 4) / 10`
- bottom margin:
  - `(PageBMargin * 530 + 4) / 10`
- left margin:
  - `(PageLMargin * 567 + 4) / 10`
- right margin:
  - `(PageRMargin * 567 + 4) / 10`

The asymmetry between vertical and horizontal constants is directly present in the decompiled code and should be preserved exactly in a faithful reconstruction.

## Grid Lines, Cell Spans, And Formatting Mapping

This is the most important section if the goal is code recreation.

### Grid line creation

Grid lines are not derived from Word constructs. They are derived from export object edges.

For every object the matrix stores:

- left edge
- right edge
- top edge
- bottom edge

These edges are inserted into `BinaryTree` structures.

After closing the trees:

- `xPos.Nodes[i].value` is the exported X grid coordinate
- `yPos.Nodes[i].value` is the exported Y grid coordinate

### Cell spans

Each object ends up with:

- `x`
- `y`
- `dx`
- `dy`

Interpretation:

- top-left cell index is `(x, y)`
- horizontal span is `dx`
- vertical span is `dy`

This is almost certainly the information used by `OoXMLDocument.Export_TableRow(...)` to decide:

- whether a cell starts here
- whether it continues from left or above
- whether `gridSpan` or vertical merge markers must be emitted

### Border mapping

Border mapping is style-driven, not object-driven at write time.

The final border visible on a cell may have been influenced by:

- original report object border
- `CloneFrames(...)` after a split
- `CheckFrameFromBottmObject(...)` during render
- `OptimizeFrames()` deduplication

Therefore the correct reconstruction rule is:

- never derive borders again from original report objects during OOXML writing
- always write final borders from the post-`Prepare()` `ExportIEMStyle`

### Fill mapping

Final cell fill comes from `ExportIEMStyle.Fill`.

But that fill may be:

- original object fill
- inherited underlying fill
- transparent base fill
- band fill

Again, write-time logic should trust the normalized style, not recompute from source objects.

### Text formatting mapping

The final writer should read text formatting from:

- `ExportIEMStyle.Font`
- `TextFill`
- `HAlign`
- `VAlign`
- `Padding`
- `LineHeight`
- `ParagraphOffset`
- `FirstTabOffset`
- `WordWrap`
- `RTL`
- `Angle`
- underline/bold/italic via the `Font`

Strong inference:

- run formatting is mostly a direct map from `Font` and `TextFill`
- paragraph formatting is mostly a map from alignment, RTL, tabs, line height, paragraph offset, and padding

## OOXML Body Writing

The only directly confirmed body writer entry point currently available is:

- `OoXMLDocument.ExportSinglePage(ExportMatrix FMatrix, bool IsLastPage, Stream file, bool ParagraphBased, bool isHeaderFooter, bool disableSectionBreak)`

### Confirmed matrix-mode writer skeleton

If `ParagraphBased == false`:

1. write `<w:tbl>`
2. call `Export_TableProperties(file, FMatrix)`
3. call `Export_TableGrid(FMatrix, file)`
4. for each row `y` from `0` to `FMatrix.Height - 2`
   - call `Export_TableRow(FMatrix, y, file, ParagraphBased, null)`
5. write `</w:tbl>`

After body content:

- if not header/footer:
  - compute page state from matrix page metadata
  - if not last page and not paragraph-based:
    - emit explicit page break paragraph
  - emit section properties either:
    - inside a trailing paragraph for non-last pages
    - or directly as terminal `<w:sectPr>` for last page

### Reconstructing `Export_TableGrid(...)`

Strongly inferred behavior:

- iterate X intervals, not X nodes
- for each column width:
  - emit a Word grid column width derived from `xPos[i+1] - xPos[i]`

### Reconstructing `Export_TableRow(...)`

Strongly inferred responsibilities:

- compute row height from `yPos[y+1] - yPos[y]`
- emit `<w:tr>`
- decide whether to emit each cell start or skip merged continuations
- resolve occupying object for each cell
- apply cell properties:
  - fill
  - borders
  - width / span
  - vertical merge
  - vertical alignment
  - margins or internal paragraph spacing
- write text/image content when the cell belongs to the owning fragment

## Header, Footer, And Watermark Writing

### Header/footer

Confirmed behavior from `OoXMLHeader.Export(...)` and `OoXMLFooter.Export(...)`:

- open `w:hdr` / `w:ftr`
- in matrix mode:
  - retrieve the corresponding header/footer matrix
  - null out hyperlinks on all objects
  - call `OoXMLDocument.ExportSinglePage(FMatrix, false, file, ParagraphBased, true, false)`
- export relationships
- close the XML part

Important:

- header/footer matrices are exported with `isHeaderFooter = true`
- therefore no section properties are appended by `ExportSinglePage(...)`

### Watermark

Confirmed behavior:

- watermark is emitted in header XML, not body XML
- text watermark uses VML shape markup
- image watermark also uses VML markup and a related media part

This means watermark support is not part of the matrix table body writer.

## Reconstruction Strategy

If the objective is to recreate the library exporter, the implementation order should be:

1. recreate `ExportIEMStyle`
2. recreate `ExportIEMObject`
3. recreate `BinaryTree` and `BinaryTreeNode`
4. recreate `ExportMatrix` exactly enough to match:
   - collection
   - snapping
   - raster fill
   - rectangle finding
   - splitting
   - border optimization
5. recreate page metadata storage equivalent to `ExportIEMPage`
6. recreate table writer:
   - table properties
   - table grid
   - table rows
   - merged cells
   - text/image emission
7. recreate package parts and relationships
8. add header/footer/watermark pipeline

Anything else is reverse-engineering in the wrong order.

## Missing Pieces For Exact Recreation

The current document is sufficient to recreate the exporter architecture and most of its logic. It is not sufficient for guaranteed source-level equivalence.

Still missing:

- full `OOExportBase`
  - needed for base export lifecycle and quoting/utilities context
- full `OoXMLDocument`
  - especially:
    - `Export_TableProperties(...)`
    - `Export_TableGrid(...)`
    - `Export_TableRow(...)`
    - page state comparison logic
    - `GetPageSize(...)`
- `ExportIEMPage`
  - needed to confirm exact page metadata storage and page break values
- `BinaryTreeNode`
  - needed for a literal recreation, though the behavior is already clear
- full image writer path
  - relation generation
  - per-cell image emission
- full paragraph/tab writer path
  - exact XML for tabs, line spacing, padding emulation, RTL, rich/html text

## What Is Already Strong Enough To Rebuild

Even with the missing pieces above, the currently known code is strong enough to rebuild:

- exporter object model
- page/band/object collection flow
- matrix normalization flow
- geometric spacing logic
- grid construction logic
- rectangle splitting logic
- border propagation and deduplication logic
- page size and margin conversion
- header/footer matrix routing
- watermark placement strategy

## Final Assessment

The library DOCX exporter is fundamentally a:

- `PreparedPage -> normalized matrix -> Word table document`

pipeline.

It is not fundamentally:

- `PreparedPage -> text boxes`

and it is not fundamentally:

- `PreparedPage -> paragraphs`

The matrix is the exporter.

If you want to recreate the code, the critical invariants to preserve are:

1. collect the same export objects
2. derive the same coordinate edges
3. snap with the same tolerance semantics
4. fill the same occupancy matrix
5. split objects the same way
6. write the final result from normalized styles and spans, not raw source objects

If more decompiled pieces become available, the next highest-value targets are:

1. `OoXMLDocument.Export_TableRow(...)`
2. `OoXMLDocument.Export_TableGrid(...)`
3. `OoXMLDocument.Export_TableProperties(...)`
4. `ExportIEMPage`
5. `BinaryTreeNode`

Those would turn this from a high-confidence reconstruction spec into an almost complete source recreation guide.

## Reconstruction Pseudocode

The pseudocode below is not guessed architecture. It is a direct normalization of the confirmed control flow and matrix behavior into implementable form.

### `Word2007Export.Start()`

```csharp
void Start()
{
    base.Start();

    Zip = new ZipArchive();
    coreDocProp = new OoXMLCoreDocumentProperties();
    applicationProp = new OoXMLApplicationProperties();
    document = new OoXMLDocument(this);
    fontTable = new OoXMLFontTable();
    wordStyles = new OoXMLWordStyles();
    wordSettings = new OoXMLWordSettings();

    document.AddRelation(1, wordSettings);
    document.AddRelation(2, wordStyles);
    document.AddRelation(3, fontTable);

    matrixList = new List<ExportMatrix>();
    pageNo = 0;
    HasWatermark = Report.PreparedPages.GetPage(0).Watermark.Enabled;

    if (paragraphBased || HasWatermark || UseHeaderAndFooter)
    {
        headers = new Dictionary<int, OoXMLHeader>();
        footers = new Dictionary<int, OoXMLFooter>();

        if (paragraphBased)
        {
            footNotes = new OoXMLFootNotes();
            endNotes = new OoXMLEndNotes();
            document.AddRelation(4, footNotes);
            document.AddRelation(5, endNotes);
        }

        int nextRelationId = paragraphBased ? 6 : 6;
        ScanPreparedPagesAndCreateHeaderFooterParts(ref nextRelationId);

        if (HasWatermark && headers.Count == 0)
        {
            headers.Add(0, new OoXMLHeader(1));
            document.AddRelation(nextRelationId++, headers[0]);
        }

        headerMatrices = new Dictionary<int, ExportMatrix>();
        footerMatrices = MatrixBased ? new Dictionary<int, ExportMatrix>() : null;
    }
}
```

### `Word2007Export.ExportPageBegin(page)`

```csharp
void ExportPageBegin(ReportPage page)
{
    base.ExportPageBegin(page);

    if (!MatrixBased)
    {
        document.ExportPageBegin(...);
        return;
    }

    headerHeight = 0;
    matrix = new ExportMatrix();
    matrix.Inaccuracy = !wysiwyg ? 10f : (paragraphBased ? 5f : 0.3f);
    if (PrintOptimized)
        matrix.ImageResolution = 300;

    matrix.RowHeightIs = rowHeightIs;
    matrix.SaveRowHeight = saveRowHeight;
    matrix.PlainRich = true;
    matrix.AreaFill = true;
    matrix.CropAreaFill = false;
    matrix.Report = Report;
    matrix.Images = true;
    matrix.WrapText = false;
    matrix.Bookmarks = Bookmarks;
    matrix.FullTrust = false;
    matrix.KeepRichText = paragraphBased;
    matrix.PrintOptimized = printOptimized;
    matrix.AddPageBegin(page);

    if (paragraphBased || HasWatermark || UseHeaderAndFooter)
    {
        headerMatrices[pageNo] = new ExportMatrix();
        footerMatrices[pageNo] = new ExportMatrix();
        ConfigureMatrix(headerMatrices[pageNo]);
        ConfigureMatrix(footerMatrices[pageNo]);
        headerMatrices[pageNo].AddPageBegin(page);
        footerMatrices[pageNo].AddPageBegin(page);
    }
}
```

### `ExportMatrix.AddBand(...)`

```csharp
void AddBand(BandBase band, object sender, bool isHeader = false)
{
    if (dataOnly && !(band is DataBand))
        return;

    if (seamless && band is PageHeaderBand && !firstPage)
        return;

    if (seamless && band is PageFooterBand)
        return;

    if (!(band.Fill is TextureFill))
        AddBandObject(band);
    else if (images)
        AddPictureObjectOrSafe(band);

    foreach (Base item in band.ForEachAllConvectedObjects(sender))
    {
        if (item is not ReportComponentBase obj || !obj.Exportable)
            continue;

        if (obj is CellularTextObject cellText)
            obj = (ReportComponentBase)cellText.GetTable();

        if (obj is TableCell)
            continue;

        if (ShouldSkipByMode(obj))
            continue;

        if (obj is TableBase table)
            AddTableObject(table);
        else if (IsMemo(obj))
            AddTextObject((TextObject)obj, isHeader);
        else if (obj is BandBase nestedBand)
            AddBandObject(nestedBand);
        else if (IsLine(obj) || (IsRect(obj) && obj.Fill.IsTransparent))
            AddLineObject(obj);
        else if (obj is HtmlObject html && htmlMode)
            AddHtml(html);
        else if (keepRichText && obj is RichObject rich)
            AddRich(rich);
        else if (images)
            AddPictureObjectOrSafe(obj);
    }
}
```

### `ExportMatrix.Prepare()`

```csharp
void Prepare()
{
    if (fillArea)
        InsertTransparentBaseObject();

    Render();
    Analyze();

    if (optFrames)
        OptimizeFrames();
}
```

### `ExportMatrix.Render()`

```csharp
void Render()
{
    xPos.Close();
    yPos.Close();

    width = xPos.Count;
    height = yPos.Count;
    matrix = new int[width * height];
    Fill(matrix, -1);

    for (int i = 0; i < iemObjectList.Count; i++)
    {
        var obj = iemObjectList[i];

        obj.x = xPos.IndexOf(obj.Left);
        int x2 = xPos.IndexOf(obj.Left + obj.Width);
        obj.dx = x2 - obj.x;

        obj.y = yPos.IndexOf(obj.Top);
        int y2 = yPos.IndexOf(obj.Top + obj.Height);
        obj.dy = y2 - obj.y;

        obj.Left = xPos.Nodes[obj.x].value;
        obj.Top = yPos.Nodes[obj.y].value;

        if (obj.Style != null && obj.Style.FillColor.A == 0)
            ApplyUnderlyingFillInheritance(obj);

        int existing = matrix[width * obj.y + obj.x];
        if (existing != -1 && obj.Style != null && iemObjectList[existing].Style != null)
            BorrowMatchingBordersFromUnderlyingObject(obj, iemObjectList[existing]);

        if (!IsPureInvisibleWhiteBand(obj))
            FillArea(obj.x, obj.y, obj.dx, obj.dy, i);
    }
}
```

### `ExportMatrix.Analyze()`

```csharp
void Analyze()
{
    for (int y = 0; y < height - 1; y++)
    {
        for (int x = 0; x < width - 1; x++)
        {
            int objIndex = matrix[width * y + x];
            if (objIndex == -1)
                continue;

            var obj = iemObjectList[objIndex];
            if (obj.Exist)
                continue;

            FindRect(x, y, out int dx, out int dy);
            ClampRectToGrid(ref dx, ref dy, x, y);

            if (obj.x != x || obj.y != y || obj.dx != dx || obj.dy != dy)
                Cut(objIndex, x, y, dx, dy);
            else
                obj.Exist = true;
        }
    }
}
```

### `OoXMLDocument.ExportSinglePage(...)`

```csharp
void ExportSinglePage(ExportMatrix matrix, bool isLastPage, Stream file,
    bool paragraphBased, bool isHeaderFooter, bool disableSectionBreak)
{
    if (!paragraphBased)
    {
        Write("<w:tbl>");
        Export_TableProperties(file, matrix);
        Export_TableGrid(matrix, file);

        for (int y = 0; y < matrix.Height - 1; y++)
            Export_TableRow(matrix, y, file, paragraphBased, null);

        Write("</w:tbl>");
    }

    if (isHeaderFooter)
        return;

    var state = new page_state(
        matrix.Landscape(0),
        ToWordWidth(matrix.PageWidth(0)),
        ToWordHeight(matrix.PageHeight(0)),
        ToWordTopMargin(matrix.PageTMargin(0)),
        ToWordBottomMargin(matrix.PageBMargin(0)),
        ToWordLeftMargin(matrix.PageLMargin(0)),
        ToWordRightMargin(matrix.PageRMargin(0)),
        paragraphBased,
        isHeaderFooter);

    if (!isLastPage && !paragraphBased)
        WriteExplicitPageBreak();

    string pageSizeXml = GetPageSize(state);
    if (string.IsNullOrEmpty(pageSizeXml))
        return;

    if (!isLastPage)
    {
        if (!disableSectionBreak)
            WriteParagraphWrappedSectionProperties(pageSizeXml);
    }
    else
    {
        WriteTerminalSectionProperties(pageSizeXml);
    }
}
```

## Exact Missing Pieces

If the target is "I can recreate the code", the missing pieces are now very specific.

### Missing and still important

1. `OoXMLDocument.Export_TableRow(...)`
   - highest-value missing function
   - this decides exact merge behavior, row height XML, cell property XML, and paragraph/image placement
2. `OoXMLDocument.Export_TableGrid(...)`
   - needed for exact width conversion and column emission
3. `OoXMLDocument.Export_TableProperties(...)`
   - needed for exact table-level layout options
4. `OoXMLDocument.GetPageSize(...)`
   - needed for exact section property serialization and change detection
5. `ExportIEMPage`
   - needed to verify exact stored page break values and page metadata field names
6. `BinaryTreeNode`
   - low conceptual risk, but still needed for literal class recreation
7. image-writing helpers
   - needed for exact `rId` handling and emitted XML
8. paragraph/text writer helpers
   - needed for exact tabs, spacing, rich text, and RTL emission

### Not missing anymore at the architecture level

These parts are already sufficiently known to rebuild confidently:

- exporter lifecycle
- matrix lifecycle
- object classification
- geometry capture
- style capture
- occupancy rasterization
- object splitting
- border propagation
- border deduplication
- header/footer routing
- watermark strategy

## Inferred Pseudocode For Missing Pieces

The pseudocode in this section is different from the previous section.

Previous pseudocode was based on directly confirmed control flow.

This section is for the still-missing library pieces. It is intentionally marked as:

- inferred
- reconstruction-oriented
- structurally likely, but not textually confirmed

The goal is to make implementation possible before the remaining decompiled code is available.

### Inferred `OoXMLDocument.Export_TableGrid(...)`

What this function almost certainly does:

- iterate the final X intervals of the matrix
- convert each interval width into Word table grid width units
- emit one `<w:gridCol>` per interval

Pseudocode:

```csharp
void Export_TableGrid(ExportMatrix matrix, Stream file)
{
    Write(file, "<w:tblGrid>");

    for (int x = 0; x < matrix.Width - 1; x++)
    {
        float left = matrix.XPosById(x);
        float right = matrix.XPosById(x + 1);
        float width = right - left;

        long wordWidth = ToWordTableWidth(width);
        Write(file, $"<w:gridCol w:w=\"{wordWidth}\"/>");
    }

    Write(file, "</w:tblGrid>");
}
```

Likely conversion rule:

- use the same coordinate basis as the rest of the document writer
- widths should be derived from matrix units after snapping, not from original report objects

### Inferred `OoXMLDocument.Export_TableProperties(...)`

What this function likely needs to define:

- table width
- borders at table level if required
- spacing/cell margins defaults
- fixed layout mode
- no accidental autofit

Pseudocode:

```csharp
void Export_TableProperties(Stream file, ExportMatrix matrix)
{
    Write(file, "<w:tblPr>");

    Write(file, "<w:tblLayout w:type=\"fixed\"/>");

    long totalWidth = 0;
    for (int x = 0; x < matrix.Width - 1; x++)
        totalWidth += ToWordTableWidth(matrix.XPosById(x + 1) - matrix.XPosById(x));

    Write(file, $"<w:tblW w:w=\"{totalWidth}\" w:type=\"dxa\"/>");

    Write(file, "<w:tblCellSpacing w:w=\"0\" w:type=\"dxa\"/>");
    Write(file, "<w:tblInd w:w=\"0\" w:type=\"dxa\"/>");

    // Likely either explicit nil borders or a neutral default border set.
    Write(file, "<w:tblBorders>");
    Write(file, "<w:top w:val=\"nil\"/>");
    Write(file, "<w:left w:val=\"nil\"/>");
    Write(file, "<w:bottom w:val=\"nil\"/>");
    Write(file, "<w:right w:val=\"nil\"/>");
    Write(file, "<w:insideH w:val=\"nil\"/>");
    Write(file, "<w:insideV w:val=\"nil\"/>");
    Write(file, "</w:tblBorders>");

    Write(file, "</w:tblPr>");
}
```

Why `fixed` layout is likely:

- the exporter is WYSIWYG-oriented
- autofit would let Word recalculate widths and destroy matrix fidelity

### Inferred `OoXMLDocument.Export_TableRow(...)`

This is the most important missing function.

Its likely inputs are:

- current matrix row index `y`
- current row height from `yPos[y+1] - yPos[y]`
- occupancy cells for all columns in that row
- merged object spans
- final `ExportIEMStyle` for each cell owner

Its likely responsibilities are:

1. emit row properties
2. determine which cells start in this row
3. suppress continuation-only positions
4. emit horizontal spans
5. emit vertical merge markers
6. emit cell formatting
7. emit text/image payload only for owning fragments

Pseudocode:

```csharp
void Export_TableRow(ExportMatrix matrix, int y, Stream file, bool paragraphBased, string styleId)
{
    float rowHeight = matrix.YPosById(y + 1) - matrix.YPosById(y);

    Write(file, "<w:tr>");
    WriteRowProperties(file, rowHeight, matrix);

    for (int x = 0; x < matrix.Width - 1; x++)
    {
        ExportIEMObject obj = matrix.GetObject(x, y);

        if (obj == null)
        {
            WriteEmptyCell(file, matrix, x, y);
            continue;
        }

        bool startsHere = obj.x == x && obj.y == y;
        bool verticalContinuation = obj.x == x && obj.y < y && y < obj.y + obj.dy;
        bool horizontalContinuation = obj.y == y && obj.x < x && x < obj.x + obj.dx;

        if (horizontalContinuation)
            continue;

        Write(file, "<w:tc>");
        WriteCellProperties(file, obj, matrix, x, y, verticalContinuation);

        if (startsHere || verticalContinuation)
        {
            if (obj.IsText)
                ExportCellText(file, obj, paragraphBased);
            else
                ExportCellGraphic(file, obj);
        }
        else
        {
            WriteEmptyParagraph(file);
        }

        Write(file, "</w:tc>");
    }

    Write(file, "</w:tr>");
}
```

### Inferred `WriteRowProperties(...)`

This is likely where `SaveRowHeight` and `RowHeightIs` matter.

Pseudocode:

```csharp
void WriteRowProperties(Stream file, float rowHeight, ExportMatrix matrix)
{
    Write(file, "<w:trPr>");

    if (matrix.SaveRowHeight)
    {
        long h = ToWordTableHeight(rowHeight);
        string rule = matrix.RowHeightIs == "min" ? "atLeast" : "exact";
        Write(file, $"<w:trHeight w:val=\"{h}\" w:hRule=\"{rule}\"/>");
    }

    Write(file, "</w:trPr>");
}
```

### Inferred `WriteCellProperties(...)`

This is the likely place where spans, merges, fill, borders, width, margins, and vertical alignment are serialized.

Pseudocode:

```csharp
void WriteCellProperties(Stream file, ExportIEMObject obj, ExportMatrix matrix, int x, int y, bool verticalContinuation)
{
    Write(file, "<w:tcPr>");

    long width = 0;
    for (int i = x; i < x + Math.Max(1, obj.dx); i++)
        width += ToWordTableWidth(matrix.XPosById(i + 1) - matrix.XPosById(i));

    Write(file, $"<w:tcW w:w=\"{width}\" w:type=\"dxa\"/>");

    if (obj.dx > 1)
        Write(file, $"<w:gridSpan w:val=\"{obj.dx}\"/>");

    if (obj.dy > 1)
    {
        string mergeState = verticalContinuation ? "continue" : "restart";
        Write(file, $"<w:vMerge w:val=\"{mergeState}\"/>");
    }

    if (obj.Style != null)
    {
        WriteVerticalAlignment(file, obj.Style.VAlign);
        WriteCellShading(file, obj.Style.FillColor);
        WriteCellBorders(file, obj.Style.Border);
        WriteCellPadding(file, obj.Style.Padding);
        WriteTextDirectionIfNeeded(file, obj.Style.Angle, obj.Style.RTL);
    }

    Write(file, "</w:tcPr>");
}
```

### Inferred text export inside a cell

There are two likely writer layers:

- paragraph properties
- run properties

Pseudocode:

```csharp
void ExportCellText(Stream file, ExportIEMObject obj, bool paragraphBased)
{
    Write(file, "<w:p>");
    WriteParagraphProperties(file, obj);

    if (obj.Hyperlink != null && obj.Hyperlink.Kind == HyperlinkKind.URL)
        WriteHyperlinkRunContainerStart(file, obj.Hyperlink);

    WriteTextRuns(file, obj);

    if (obj.Hyperlink != null && obj.Hyperlink.Kind == HyperlinkKind.URL)
        WriteHyperlinkRunContainerEnd(file);

    Write(file, "</w:p>");
}
```

### Inferred `WriteParagraphProperties(...)`

Likely sources:

- `HAlign`
- `RTL`
- `Padding`
- `LineHeight`
- `ParagraphOffset`
- `TabPositions`
- `FirstTabOffset`
- maybe `WordWrap`

Pseudocode:

```csharp
void WriteParagraphProperties(Stream file, ExportIEMObject obj)
{
    ExportIEMStyle style = obj.Style;

    Write(file, "<w:pPr>");

    WriteParagraphAlignment(file, style.HAlign, style.RTL, style.ForceJustify);
    WriteParagraphSpacing(file, style.LineHeight, style.ParagraphOffset);
    WriteParagraphIndent(file, style.Padding, style.FirstTabOffset);
    WriteParagraphTabs(file, obj.TabPositions, obj.TabWidth, style.FirstTabOffset);

    if (style.RTL)
        Write(file, "<w:bidi/>");

    Write(file, "</w:pPr>");
}
```

Important note:

- `Padding.Top` and `Padding.Bottom` might not map to real Word paragraph spacing directly
- they may instead be emulated with cell margins or with a combination of paragraph spacing and row height
- this is one of the places where exact output still depends on the missing original code

### Inferred `WriteTextRuns(...)`

Likely responsibilities:

- split on line breaks
- preserve tabs
- preserve HTML/rich text if specialized path exists
- use font/color formatting

Pseudocode:

```csharp
void WriteTextRuns(Stream file, ExportIEMObject obj)
{
    if (obj.IsRichText)
    {
        WriteRichTextPayload(file, obj);
        return;
    }

    IEnumerable<TextToken> tokens = TokenizeText(obj.Text, obj.TextRenderType);

    foreach (var token in tokens)
    {
        switch (token.Kind)
        {
            case TextTokenKind.LineBreak:
                Write(file, "<w:r><w:br/></w:r>");
                break;

            case TextTokenKind.Tab:
                Write(file, "<w:r><w:tab/></w:r>");
                break;

            case TextTokenKind.Text:
                Write(file, "<w:r>");
                WriteRunProperties(file, obj.Style);
                WriteEscapedText(file, token.Text);
                Write(file, "</w:r>");
                break;
        }
    }
}
```

### Inferred `WriteRunProperties(...)`

Pseudocode:

```csharp
void WriteRunProperties(Stream file, ExportIEMStyle style)
{
    Write(file, "<w:rPr>");

    WriteFontName(file, style.Font.Name);
    WriteFontSize(file, style.Font.SizeInPoints);
    WriteTextColor(file, style.TextColor);

    if (style.Font.Bold)
        Write(file, "<w:b/>");
    if (style.Font.Italic)
        Write(file, "<w:i/>");
    if (style.Font.Underline || style.Font.Underlines)
        Write(file, "<w:u w:val=\"single\"/>");
    if (style.Font.Strikeout)
        Write(file, "<w:strike/>");

    if (style.Angle != 0)
        WriteRotatedTextHints(file, style.Angle);

    Write(file, "</w:rPr>");
}
```

### Inferred image export

The matrix object already carries:

- `PictureStream`
- `Metafile`
- `Hash`
- `Base`
- maybe `ImageData`

So the likely writer behavior is:

1. ensure media part exists or is reused by hash
2. create/get relationship id
3. emit drawing or pict markup inside the owning cell paragraph

Pseudocode:

```csharp
void ExportCellGraphic(Stream file, ExportIEMObject obj)
{
    if (obj.PictureStream == null && obj.Metafile == null && !obj.IsSvg)
    {
        WriteEmptyParagraph(file);
        return;
    }

    string relationId = EnsureImageRelation(obj);

    Write(file, "<w:p>");
    Write(file, "<w:r>");
    WriteInlineOrAnchoredImage(file, relationId, obj.Width, obj.Height, obj.PicBounds);
    Write(file, "</w:r>");
    Write(file, "</w:p>");
}
```

### Inferred `GetPageSize(...)`

The visible code shows this function returns XML only when the current page state differs from the previously emitted one, or when a section must be declared.

Likely responsibilities:

- compare page orientation, width, height, margins, paragraph mode, header/footer state
- emit `<w:pgSz .../>`
- emit `<w:pgMar .../>`
- emit header/footer references where applicable

Pseudocode:

```csharp
string GetPageSize(page_state state)
{
    if (PageStateEquals(lastState, state))
        return string.Empty;

    var sb = new StringBuilder();

    sb.Append($"<w:pgSz w:w=\"{state.Width}\" w:h=\"{state.Height}\"");
    if (state.Landscape)
        sb.Append(" w:orient=\"landscape\"");
    sb.Append("/>");

    sb.Append($"<w:pgMar w:top=\"{state.TopMargin}\" w:right=\"{state.RightMargin}\" " +
              $"w:bottom=\"{state.BottomMargin}\" w:left=\"{state.LeftMargin}\"/>");

    AppendHeaderFooterReferencesIfNeeded(sb, state);

    lastState = state;
    return sb.ToString();
}
```

### Inferred `ExportIEMPage`

Based on what `ExportMatrix` reads back, `ExportIEMPage` likely contains at least:

- `Landscape`
- `Width`
- `Height`
- `RawPaperSize`
- `LeftMargin`
- `TopMargin`
- `RightMargin`
- `BottomMargin`
- `Value` as page break coordinate
- maybe watermark stream storage

Pseudocode:

```csharp
class ExportIEMPage
{
    public bool Landscape;
    public float Width;
    public float Height;
    public int RawPaperSize;
    public float LeftMargin;
    public float TopMargin;
    public float RightMargin;
    public float BottomMargin;
    public float Value;
    public MemoryStream WatermarkPictureStream;
}
```

### Inferred `BinaryTreeNode`

Its structure is strongly implied by `BinaryTree`.

Pseudocode:

```csharp
class BinaryTreeNode
{
    public float value;
    public int index;
    public BinaryTreeNode left;
    public BinaryTreeNode right;
    public int leftCount;
    public int rightCount;

    public BinaryTreeNode(float value)
    {
        this.value = value;
    }
}
```
