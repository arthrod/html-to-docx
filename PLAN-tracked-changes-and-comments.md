# Implementation Plan: HTML Tracked Changes & Comments to DOCX

## Overview

Add support for converting HTML `<ins>`, `<del>`, and comment markup into proper OpenXML tracked changes and threaded comments in generated DOCX files. Currently, `<ins>` renders as underline and `<del>` as strikethrough — purely visual. This plan converts them into real Word revision markup (`w:ins`, `w:del`) and real Word comments (`comments.xml` + `commentsExtended.xml`).

---

## 1. HTML Input Format (Proposed Conventions)

### 1.1 Insertions — `<ins>`

Uses the standard HTML `<ins>` element with `data-*` attributes for OpenXML metadata.

```html
<p>
  The project
  <ins data-author="Jane Doe" datetime="2024-06-15T10:30:00Z" data-id="1">
    has been completed
  </ins>
  on schedule.
</p>
```

| HTML Attribute | OpenXML Attribute | Required | Default |
|----------------|-------------------|----------|---------|
| `data-author`  | `w:author`        | No       | `"Unknown Author"` |
| `datetime`     | `w:date`          | No       | current ISO timestamp |
| `data-id`      | `w:id`            | No       | auto-incremented integer |

### 1.2 Deletions — `<del>`

Uses the standard HTML `<del>` element. Text inside `<del>` maps to `w:delText` (not `w:t`).

```html
<p>
  Some
  <del data-author="Joe Smith" datetime="2024-06-15T12:50:00Z" data-id="2">
    removed
  </del>
  text.
</p>
```

| HTML Attribute | OpenXML Attribute | Required | Default |
|----------------|-------------------|----------|---------|
| `data-author`  | `w:author`        | No       | `"Unknown Author"` |
| `datetime`     | `w:date`          | No       | current ISO timestamp |
| `data-id`      | `w:id`            | No       | auto-incremented integer |

### 1.3 Comments — `<span class="comment-start">` / `<span class="comment-end">`

Following Pandoc's convention for comment range markers (the most widely-used open-source tool that handles DOCX comments):

```html
<p>
  This is
  <span class="comment-start" id="0"
        data-author="Alice Johnson"
        data-initials="AJ"
        data-date="2024-06-15T10:00:00Z"
        data-comment="This paragraph needs revision for clarity."></span>
  the commented text
  <span class="comment-end" id="0"></span>
  in the document.
</p>
```

| HTML Attribute  | OpenXML Mapping | Required | Default |
|-----------------|-----------------|----------|---------|
| `id`            | `w:id` on `commentRangeStart`/`End`/`commentReference` + `w:comment` | **Yes** | — |
| `data-author`   | `w:author` on `w:comment` | No | `"Unknown Author"` |
| `data-initials` | `w:initials` on `w:comment` | No | derived from author |
| `data-date`     | `w:date` on `w:comment` | No | current ISO timestamp |
| `data-comment`  | Text body of the `w:comment` element | **Yes** | — |

### 1.4 Threaded Comments (Replies)

Replies are additional `<span class="comment-start">` elements with a `data-parent-id` linking to the parent comment:

```html
<span class="comment-start" id="1"
      data-author="Bob Williams"
      data-date="2024-06-15T11:30:00Z"
      data-comment="I agree, I will update it."
      data-parent-id="0"></span>
<span class="comment-end" id="1"></span>
```

| HTML Attribute    | OpenXML Mapping | Required | Default |
|-------------------|-----------------|----------|---------|
| `data-parent-id`  | `w15:paraIdParent` in `commentsExtended.xml` | No | top-level comment |
| `data-done`       | `w15:done` in `commentsExtended.xml` | No | `"0"` (open) |

### 1.5 Combined Example

```html
<p>
  The project
  <del data-author="Editor" datetime="2024-06-15T09:00:00Z">will be</del>
  <ins data-author="Editor" datetime="2024-06-15T09:00:00Z">has been</ins>
  <span class="comment-start" id="0"
        data-author="Reviewer"
        data-date="2024-06-15T10:00:00Z"
        data-comment="Consider rephrasing this section."></span>
  completed on schedule
  <span class="comment-end" id="0"></span>
  and delivered to the client.
</p>
```

---

## 2. OpenXML Output Structures

### 2.1 Track Changes: Insertions in `document.xml`

```xml
<w:ins w:id="1" w:author="Jane Doe" w:date="2024-06-15T10:30:00Z">
  <w:r>
    <w:t>has been completed</w:t>
  </w:r>
</w:ins>
```

### 2.2 Track Changes: Deletions in `document.xml`

Note: deleted text MUST use `<w:delText>`, not `<w:t>`.

```xml
<w:del w:id="2" w:author="Joe Smith" w:date="2024-06-15T12:50:00Z">
  <w:r>
    <w:delText xml:space="preserve">removed</w:delText>
  </w:r>
</w:del>
```

### 2.3 Comments in `document.xml`

Three elements anchor a comment to a text range:

```xml
<!-- Start of commented range -->
<w:commentRangeStart w:id="0"/>

<w:r>
  <w:t>the commented text</w:t>
</w:r>

<!-- End of commented range -->
<w:commentRangeEnd w:id="0"/>

<!-- Comment reference marker (the superscript indicator) -->
<w:r>
  <w:rPr>
    <w:rStyle w:val="CommentReference"/>
  </w:rPr>
  <w:commentReference w:id="0"/>
</w:r>
```

### 2.4 `word/comments.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:comment w:id="0" w:author="Alice Johnson" w:initials="AJ"
             w:date="2024-06-15T10:00:00Z">
    <w:p w14:paraId="1A2B3C01" w14:textId="5E6F7A01">
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:annotationRef/>
      </w:r>
      <w:r>
        <w:t>This paragraph needs revision for clarity.</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>
```

### 2.5 `word/commentsExtended.xml` (Threaded Comments)

Links via `w14:paraId` on the **last paragraph** in each `<w:comment>`.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="w15">
  <!-- Top-level comment -->
  <w15:commentEx w15:paraId="1A2B3C01" w15:done="0"/>
  <!-- Reply to the above -->
  <w15:commentEx w15:paraId="9C8D7E01" w15:paraIdParent="1A2B3C01" w15:done="0"/>
</w15:commentsEx>
```

### 2.6 `word/settings.xml` — Enable Track Changes

```xml
<w:settings ...>
  ...
  <w:trackRevisions w:val="true"/>
  ...
</w:settings>
```

---

## 3. New OpenXML Parts & Metadata

### 3.1 New Files in DOCX ZIP

| File | When Added |
|------|------------|
| `word/comments.xml` | When any `<span class="comment-start">` is present |
| `word/commentsExtended.xml` | When any `data-parent-id` or `data-done` attribute is present (threaded comments) |

### 3.2 New Relationships in `word/_rels/document.xml.rels`

| Relationship Type | Target | When Added |
|-------------------|--------|------------|
| `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` | `comments.xml` | When comments present |
| `http://schemas.microsoft.com/office/2011/relationships/commentsExtended` | `commentsExtended.xml` | When threaded comments present |

### 3.3 New Content Types in `[Content_Types].xml`

| PartName | ContentType | When Added |
|----------|-------------|------------|
| `/word/comments.xml` | `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml` | When comments present |
| `/word/commentsExtended.xml` | `application/vnd.ms-word.commentsExtended+xml` | When threaded comments present |

### 3.4 New Styles in `word/styles.xml`

Two styles are required for comments to render correctly:

```xml
<w:style w:type="character" w:styleId="CommentReference">
  <w:name w:val="annotation reference"/>
  <w:basedOn w:val="DefaultParagraphFont"/>
  <w:semiHidden/>
  <w:unhideWhenUsed/>
  <w:rPr>
    <w:sz w:val="16"/>
    <w:szCs w:val="16"/>
  </w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="CommentText">
  <w:name w:val="annotation text"/>
  <w:basedOn w:val="Normal"/>
  <w:semiHidden/>
  <w:unhideWhenUsed/>
  <w:rPr>
    <w:sz w:val="20"/>
    <w:szCs w:val="20"/>
  </w:rPr>
</w:style>
```

### 3.5 New Namespaces

| Prefix | URI | Purpose |
|--------|-----|---------|
| `w14`  | `http://schemas.microsoft.com/office/word/2010/wordml` | `paraId` / `textId` on comment paragraphs |
| `w15`  | `http://schemas.microsoft.com/office/word/2012/wordml` | `commentsExtended.xml` |
| `mc`   | `http://schemas.openxmlformats.org/markup-compatibility/2006` | Markup compatibility (already exists as `ve`) |

---

## 4. File-by-File Implementation Plan

### Step 1: `src/namespaces.js` — Add New Namespace URIs

Add the following entries:

```javascript
w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
commentsExtended: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
```

**Rationale:** These are standard OOXML namespace URIs. `w14` and `w15` are needed for `paraId`/`textId` attributes on comment paragraphs and the `commentsExtended.xml` file, respectively. The relationship URIs are needed when adding entries to `document.xml.rels`.

---

### Step 2: `src/constants.js` — Add New Type Constants

Add:

```javascript
const commentType = 'comment';
const commentsExtendedType = 'commentsExtended';
const commentsFileName = 'comments';
const commentsExtendedFileName = 'commentsExtended';
const defaultTrackChangesAuthor = 'Unknown Author';
```

Export them alongside existing constants.

---

### Step 3: `src/schemas/settings.js` — Conditionally Add `trackRevisions`

**Current:** Static XML string with no `trackRevisions`.

**Change:** The settings generation must be made dynamic — if tracked changes (`<ins>` or `<del>` with `data-author`) are present in the HTML, add `<w:trackRevisions/>` inside `<w:settings>`.

**Approach:** Instead of making settings a static string, pass a flag from `DocxDocument`:

```javascript
const generateSettingsXML = (options = {}) => {
  const trackRevisionsXml = options.trackRevisions
    ? '<w:trackRevisions/>'
    : '';

  return `
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:settings xmlns:w="${namespaces.w}" xmlns:o="${namespaces.o}"
                xmlns:r="${namespaces.r}" xmlns:v="${namespaces.v}"
                xmlns:w10="${namespaces.w10}" xmlns:sl="${namespaces.sl}">
      <w:zoom w:percent="100"/>
      <w:defaultTabStop w:val="720"/>
      ${trackRevisionsXml}
      <w:decimalSymbol w:val="."/>
      <w:listSeparator w:val=","/>
    </w:settings>
  `;
};
```

---

### Step 4: `src/schemas/styles.js` — Add Comment Styles

Add `CommentReference` and `CommentText` styles when comments are present:

```javascript
// Add inside generateStylesXML, conditionally when hasComments is true:
const commentStylesXml = hasComments ? `
  <w:style w:type="character" w:styleId="CommentReference">
    <w:name w:val="annotation reference"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:sz w:val="16"/>
      <w:szCs w:val="16"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CommentText">
    <w:name w:val="annotation text"/>
    <w:basedOn w:val="Normal"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
` : '';
```

---

### Step 5: `src/schemas/comments.js` — New File

Template for `comments.xml`:

```javascript
import namespaces from '../namespaces';

const commentsXML = `
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:comments xmlns:w="${namespaces.w}"
              xmlns:w14="${namespaces.w14}"
              xmlns:r="${namespaces.r}">
  </w:comments>
`;

export default commentsXML;
```

---

### Step 6: `src/schemas/comments-extended.js` — New File

Template for `commentsExtended.xml`:

```javascript
import namespaces from '../namespaces';

const commentsExtendedXML = `
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w15:commentsEx xmlns:w15="${namespaces.w15}"
                  xmlns:mc="${namespaces.ve}"
                  mc:Ignorable="w15">
  </w15:commentsEx>
`;

export default commentsExtendedXML;
```

---

### Step 7: `src/schemas/index.js` — Export New Schemas

Add:

```javascript
export { default as commentsXML } from './comments';
export { default as commentsExtendedXML } from './comments-extended';
```

---

### Step 8: `src/schemas/document.template.js` — Add `w14` Namespace

The document template needs the `w14` namespace declared for `paraId` references to work:

```xml
xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
```

Add these to the `<w:document>` element alongside the existing namespace declarations.

---

### Step 9: `src/docx-document.js` — Core Orchestration Changes

This is the central coordination point. Changes needed:

#### 9a. New State Properties in Constructor

```javascript
// Track changes state
this.lastRevisionId = 0;        // Auto-incrementing revision ID for w:ins/w:del
this.hasTrackedChanges = false;  // Flag to enable trackRevisions in settings.xml

// Comments state
this.commentObjects = [];        // Array of { id, author, initials, date, text, paraId, textId }
this.commentExtendedObjects = []; // Array of { paraId, parentParaId, done }
this.hasComments = false;
this.hasThreadedComments = false;
this.lastCommentId = -1;         // Will be set from HTML data-id or auto-incremented
```

#### 9b. New Methods

**`createRevisionId()`** — Returns next unique revision ID (shared counter for `w:ins`, `w:del`, `commentRangeStart`, `commentRangeEnd`, `commentReference`):

```javascript
createRevisionId() {
  this.lastRevisionId += 1;
  return this.lastRevisionId;
}
```

**`addComment(commentData)`** — Registers a comment for later XML generation:

```javascript
addComment({ id, author, initials, date, text, parentId, done }) {
  this.hasComments = true;
  const paraId = this.generateParaId();
  const textId = this.generateParaId();
  const commentObj = { id, author, initials, date, text, paraId, textId };
  this.commentObjects.push(commentObj);

  // Always add to extended for done tracking; add parentParaId if threaded
  const parentComment = parentId != null
    ? this.commentObjects.find(c => c.id === parseInt(parentId))
    : null;
  if (parentComment || done) {
    this.hasThreadedComments = true;
  }
  this.commentExtendedObjects.push({
    paraId,
    parentParaId: parentComment ? parentComment.paraId : null,
    done: done || '0',
  });

  return commentObj;
}
```

**`generateParaId()`** — Generates an 8-character hex string for `w14:paraId`:

```javascript
generateParaId() {
  // Generate deterministic or random 8-char hex ID
  if (this.deterministicIds) {
    this.lastRevisionId += 1;
    return this.lastRevisionId.toString(16).toUpperCase().padStart(8, '0');
  }
  return nanoid(8).replace(/[^0-9A-Fa-f]/g, '0').toUpperCase().padStart(8, '0');
}
```

**`generateCommentsXML()`** — Builds the complete `comments.xml`:

```javascript
generateCommentsXML() {
  if (!this.hasComments) return null;
  const commentsXML = create({ encoding: 'UTF-8', standalone: true }, commentsXMLTemplate);

  this.commentObjects.forEach(({ id, author, initials, date, text, paraId, textId }) => {
    const commentFragment = fragment({ namespaceAlias: { w: namespaces.w, w14: namespaces.w14 } })
      .ele('@w', 'comment')
      .att('@w', 'id', String(id))
      .att('@w', 'author', author)
      .att('@w', 'initials', initials)
      .att('@w', 'date', date)
        .ele('@w', 'p')
        .att('@w14', 'paraId', paraId)
        .att('@w14', 'textId', textId)
          .ele('@w', 'pPr')
            .ele('@w', 'pStyle').att('@w', 'val', 'CommentText').up()
          .up()
          .ele('@w', 'r')
            .ele('@w', 'rPr')
              .ele('@w', 'rStyle').att('@w', 'val', 'CommentReference').up()
            .up()
            .ele('@w', 'annotationRef').up()
          .up()
          .ele('@w', 'r')
            .ele('@w', 't').txt(text).up()
          .up()
        .up()
      .up();

    commentsXML.root().import(commentFragment);
  });

  return commentsXML.toString({ prettyPrint: true });
}
```

**`generateCommentsExtendedXML()`** — Builds `commentsExtended.xml`:

```javascript
generateCommentsExtendedXML() {
  if (!this.hasThreadedComments && !this.hasComments) return null;
  const extXML = create({ encoding: 'UTF-8', standalone: true }, commentsExtendedXMLTemplate);

  this.commentExtendedObjects.forEach(({ paraId, parentParaId, done }) => {
    const exFragment = fragment({ namespaceAlias: { w15: namespaces.w15 } })
      .ele('@w15', 'commentEx')
      .att('@w15', 'paraId', paraId)
      .att('@w15', 'done', done);

    if (parentParaId) {
      exFragment.att('@w15', 'paraIdParent', parentParaId);
    }
    exFragment.up();

    extXML.root().import(exFragment);
  });

  return extXML.toString({ prettyPrint: true });
}
```

#### 9c. Modify `generateSettingsXML()`

Pass the `hasTrackedChanges` flag:

```javascript
generateSettingsXML() {
  return generateXMLString(generateSettingsXML({ trackRevisions: this.hasTrackedChanges }));
}
```

#### 9d. Modify `generateStylesXML()`

Pass the `hasComments` flag so comment styles are included.

#### 9e. Modify `generateContentTypesXML()`

Add content type overrides for comments files when present:

```javascript
if (this.hasComments) {
  const commentsOverride = fragment({ defaultNamespace: { ele: namespaces.contentTypes } })
    .ele('Override')
    .att('PartName', '/word/comments.xml')
    .att('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml')
    .up();
  contentTypesXML.root().import(commentsOverride);
}
if (this.hasThreadedComments) {
  const extOverride = fragment({ defaultNamespace: { ele: namespaces.contentTypes } })
    .ele('Override')
    .att('PartName', '/word/commentsExtended.xml')
    .att('ContentType', 'application/vnd.ms-word.commentsExtended+xml')
    .up();
  contentTypesXML.root().import(extOverride);
}
```

#### 9f. Modify `createDocumentRelationships()`

Add cases for the new relationship types:

```javascript
case commentType:
  relationshipType = namespaces.comments;
  break;
case commentsExtendedType:
  relationshipType = namespaces.commentsExtended;
  break;
```

---

### Step 10: `src/helpers/render-document-file.js` — Handle New HTML Elements

#### 10a. Change `<ins>` and `<del>` Routing

**Current behavior** (lines 236-239):
```javascript
case 'ins':  // falls through to paragraph builder → renders as underline
case 'del':  // falls through to paragraph builder → renders as strikethrough
```

**New behavior:**
When `<ins>` or `<del>` has a `data-author` attribute, treat it as a tracked change. Otherwise, fall through to the existing visual-only behavior (backward compatible).

```javascript
case 'ins':
case 'del':
  if (vNode.properties.attributes && vNode.properties.attributes['data-author']) {
    // Tracked change mode
    const trackedFragment = await xmlBuilder.buildTrackedChange(
      vNode,
      docxDocumentInstance
    );
    xmlFragment.import(trackedFragment);
    return;
  }
  // Fall through to existing visual-only behavior
  break;
```

Then the existing `case 'ins': case 'del':` in the paragraph-building switch handles visual rendering.

#### 10b. Handle `<span class="comment-start">` and `<span class="comment-end">`

In the `case 'span':` handling, detect comment markers:

```javascript
case 'span':
  const spanClass = vNode.properties.attributes?.class;
  if (spanClass === 'comment-start') {
    const commentFragment = xmlBuilder.buildCommentRangeStart(vNode, docxDocumentInstance);
    xmlFragment.import(commentFragment);
    return;
  }
  if (spanClass === 'comment-end') {
    const commentFragment = xmlBuilder.buildCommentRangeEnd(vNode, docxDocumentInstance);
    xmlFragment.import(commentFragment);
    return;
  }
  // Fall through to existing paragraph building
  break;
```

---

### Step 11: `src/helpers/xml-builder.js` — New Builder Functions

#### 11a. `buildTrackedChange(vNode, docxDocumentInstance)`

Wraps child content in `<w:ins>` or `<w:del>`:

```javascript
export async function buildTrackedChange(vNode, docxDocumentInstance) {
  const isInsertion = vNode.tagName === 'ins';
  const author = vNode.properties.attributes['data-author'] || 'Unknown Author';
  const date = vNode.properties.attributes.datetime || new Date().toISOString();
  const id = vNode.properties.attributes['data-id']
    || docxDocumentInstance.createRevisionId();

  docxDocumentInstance.hasTrackedChanges = true;

  const wrapperTag = isInsertion ? 'ins' : 'del';

  const changeFragment = fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', wrapperTag)
    .att('@w', 'id', String(id))
    .att('@w', 'author', author)
    .att('@w', 'date', date);

  // Build child runs inside the tracked change wrapper
  // For deletions, text elements must use w:delText instead of w:t
  const childRuns = await buildRunsFromChildren(
    vNode,
    docxDocumentInstance,
    { useDelText: !isInsertion }
  );
  changeFragment.import(childRuns);
  changeFragment.up();

  return changeFragment;
}
```

#### 11b. `buildCommentRangeStart(vNode, docxDocumentInstance)`

```javascript
export function buildCommentRangeStart(vNode, docxDocumentInstance) {
  const attrs = vNode.properties.attributes;
  const commentId = attrs.id;
  const author = attrs['data-author'] || 'Unknown Author';
  const initials = attrs['data-initials'] || deriveInitials(author);
  const date = attrs['data-date'] || new Date().toISOString();
  const text = attrs['data-comment'] || '';
  const parentId = attrs['data-parent-id'];
  const done = attrs['data-done'] || '0';

  // Register the comment with the document
  docxDocumentInstance.addComment({ id: parseInt(commentId), author, initials, date, text, parentId, done });

  // Emit commentRangeStart
  return fragment({ namespaceAlias: { w: namespaces.w } })
    .ele('@w', 'commentRangeStart')
    .att('@w', 'id', String(commentId))
    .up();
}
```

#### 11c. `buildCommentRangeEnd(vNode, docxDocumentInstance)`

```javascript
export function buildCommentRangeEnd(vNode, docxDocumentInstance) {
  const commentId = vNode.properties.attributes.id;

  const frag = fragment({ namespaceAlias: { w: namespaces.w } });

  // commentRangeEnd
  frag.import(
    fragment({ namespaceAlias: { w: namespaces.w } })
      .ele('@w', 'commentRangeEnd')
      .att('@w', 'id', String(commentId))
      .up()
  );

  // commentReference run
  frag.import(
    fragment({ namespaceAlias: { w: namespaces.w } })
      .ele('@w', 'r')
        .ele('@w', 'rPr')
          .ele('@w', 'rStyle').att('@w', 'val', 'CommentReference').up()
        .up()
        .ele('@w', 'commentReference').att('@w', 'id', String(commentId)).up()
      .up()
  );

  return frag;
}
```

#### 11d. `buildRunsFromChildren(vNode, docxDocumentInstance, options)`

A helper that builds `<w:r>` elements from `<ins>`/`<del>` children, using `<w:delText>` for deletions:

```javascript
async function buildRunsFromChildren(vNode, docxDocumentInstance, options = {}) {
  const xmlFrag = fragment({ namespaceAlias: { w: namespaces.w } });
  const { useDelText } = options;

  // Iterate child nodes, building runs
  for (const child of vNode.children || []) {
    if (isVText(child)) {
      const runFragment = fragment({ namespaceAlias: { w: namespaces.w } })
        .ele('@w', 'r');

      // Apply run properties from parent formatting
      const rPr = buildRunPropertiesFromNode(vNode, docxDocumentInstance);
      if (rPr) runFragment.import(rPr);

      const textTag = useDelText ? 'delText' : 't';
      runFragment.ele('@w', textTag)
        .att('xml:space', 'preserve')
        .txt(child.text)
        .up();

      runFragment.up();
      xmlFrag.import(runFragment);
    } else if (isVNode(child)) {
      // Handle nested inline elements (e.g., <strong> inside <ins>)
      const nestedRuns = await buildRunsFromChildren(child, docxDocumentInstance, options);
      xmlFrag.import(nestedRuns);
    }
  }

  return xmlFrag;
}
```

#### 11e. `deriveInitials(authorName)` — Utility

```javascript
function deriveInitials(authorName) {
  return authorName
    .split(/\s+/)
    .map(word => word.charAt(0).toUpperCase())
    .join('');
}
```

#### 11f. Modify Existing `buildRunProperties()` Tag Handling

In the function that maps tag names to formatting (around line 705-714), the `ins`/`del`/`strike` cases need a comment clarifying behavior:

```javascript
case 'ins':
case 'u':
  // Visual underline only (no data-author) — backward compatible
  return buildUnderline();
case 'strike':
case 'del':
case 's':
  // Visual strikethrough only (no data-author) — backward compatible
  return buildStrike();
```

The routing decision (tracked change vs. visual) happens upstream in `render-document-file.js` before these functions are called.

---

### Step 12: `src/html-to-docx.js` — Add Comments Files to ZIP

After the document XML is generated and comment objects are collected:

```javascript
// After zip.folder(wordFolder).file('document.xml', ...)

// Generate and add comments.xml if comments are present
if (docxDocument.hasComments) {
  zip.folder(wordFolder).file('comments.xml', docxDocument.generateCommentsXML(), {
    createFolders: false,
  });

  // Add relationship for comments
  docxDocument.createDocumentRelationships(
    docxDocument.relationshipFilename,
    commentType,
    'comments.xml',
    internalRelationship
  );
}

// Generate and add commentsExtended.xml if threaded comments present
if (docxDocument.hasThreadedComments || docxDocument.hasComments) {
  const extXML = docxDocument.generateCommentsExtendedXML();
  if (extXML) {
    zip.folder(wordFolder).file('commentsExtended.xml', extXML, {
      createFolders: false,
    });

    docxDocument.createDocumentRelationships(
      docxDocument.relationshipFilename,
      commentsExtendedType,
      'commentsExtended.xml',
      internalRelationship
    );
  }
}
```

**Important ordering note:** The `createDocumentRelationships` calls for comments must happen BEFORE `generateRelsXML()` is called, which is already at the end of `addFilesToContainer`. The comments file writes should be placed after `document.xml` but before the rels generation.

---

### Step 13: `src/schemas/content-types.js` — No Static Changes Needed

Content types are added dynamically in `docx-document.js` `generateContentTypesXML()` (Step 9e). No changes to the static template.

---

### Step 14: `index.js` — Preserve Comment Markers During Minification

**Critical issue:** The HTML minifier's `removeComments: true` option will strip HTML comments (`<!-- -->`), but our comment markers are `<span>` elements, so they survive minification. However, empty `<span>` elements might be removed by aggressive minification.

Update minification options to preserve empty spans:

```javascript
const minifiedHTMLString = await minify(htmlString, {
  collapseWhitespace: true,
  removeComments: true,
  removeEmptyElements: false,  // Preserve <span class="comment-start"></span>
});
```

Verify that `html-minifier-terser` does not remove empty spans by default (it does not unless `removeEmptyElements: true` is set, which is not the current config). This is safe.

---

## 5. Backward Compatibility

### Key Design Decisions

1. **`<ins>` without `data-author`** → renders as underline (existing behavior)
2. **`<del>` without `data-author`** → renders as strikethrough (existing behavior)
3. **`<ins>` with `data-author`** → renders as tracked insertion (`w:ins`)
4. **`<del>` with `data-author`** → renders as tracked deletion (`w:del`)
5. **No comments in HTML** → no `comments.xml` generated, no extra relationships
6. **`trackRevisions` in settings.xml** → only added when tracked changes are present

This ensures 100% backward compatibility: existing HTML without `data-author` attributes produces identical DOCX output.

---

## 6. Testing Plan

### 6.1 Reference DOCX Fixture

Create a real DOCX file (created in Microsoft Word or LibreOffice) that contains:
- Normal text
- An insertion (tracked change)
- A deletion (tracked change)
- A comment on a text range
- A threaded comment (reply)
- A resolved comment (done=1)
- Mixed: insertion + comment on same range

Store at `tests/fixtures/tracked-changes-and-comments.docx`.

This fixture serves as a reference for the exact XML structure expected and validates that the output is Word-compatible.

### 6.2 Unit Tests — New File: `tests/tracked-changes.test.js`

```javascript
describe('Track Changes', () => {
  describe('Insertions', () => {
    test('should generate w:ins for <ins data-author="...">', async () => {
      const html = '<p>Hello <ins data-author="Jane" datetime="2024-01-01T00:00:00Z">world</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      // Verify w:ins element exists in document.xml
      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('w:author="Jane"');
      expect(parsed.xml).toContain('<w:t');
      expect(parsed.xml).toContain('world');
    });

    test('should fall back to underline for <ins> without data-author', async () => {
      const html = '<p>Hello <ins>world</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      // Should NOT have w:ins (tracked change)
      expect(parsed.xml).not.toContain('<w:ins ');
      // Should have underline formatting
      expect(parsed.xml).toContain('<w:u ');
    });

    test('should auto-generate revision ID when data-id not provided', async () => {
      const html = '<p><ins data-author="Jane">text</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toMatch(/w:ins[^>]*w:id="\d+"/);
    });

    test('should use provided data-id for revision ID', async () => {
      const html = '<p><ins data-author="Jane" data-id="42">text</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('w:id="42"');
    });

    test('should handle nested formatting inside insertions', async () => {
      const html = '<p><ins data-author="Jane"><strong>bold inserted</strong></ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('<w:b');
    });
  });

  describe('Deletions', () => {
    test('should generate w:del with w:delText for <del data-author="...">', async () => {
      const html = '<p>Some <del data-author="Joe" datetime="2024-01-01T00:00:00Z">removed</del> text</p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:del ');
      expect(parsed.xml).toContain('w:author="Joe"');
      expect(parsed.xml).toContain('<w:delText');
      expect(parsed.xml).toContain('removed');
      // Must NOT use w:t for deleted text
      expect(parsed.xml).not.toMatch(/<w:del[^]*?<w:t[^>]*>removed/);
    });

    test('should fall back to strikethrough for <del> without data-author', async () => {
      const html = '<p>Some <del>removed</del> text</p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).not.toContain('<w:del ');
      expect(parsed.xml).toContain('<w:strike');
    });
  });

  describe('Settings', () => {
    test('should add trackRevisions to settings.xml when tracked changes present', async () => {
      const html = '<p><ins data-author="Jane">text</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const settingsXml = await zip.file('word/settings.xml').async('string');

      expect(settingsXml).toContain('<w:trackRevisions');
    });

    test('should NOT add trackRevisions when no tracked changes', async () => {
      const html = '<p>Normal text</p>';
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const settingsXml = await zip.file('word/settings.xml').async('string');

      expect(settingsXml).not.toContain('trackRevisions');
    });
  });

  describe('Mixed insertions and deletions', () => {
    test('should handle adjacent ins and del', async () => {
      const html = '<p><del data-author="Ed">old</del><ins data-author="Ed">new</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:del ');
      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('<w:delText');
    });
  });
});
```

### 6.3 Unit Tests — New File: `tests/comments.test.js`

```javascript
describe('Comments', () => {
  describe('Basic comments', () => {
    test('should generate comments.xml with comment content', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="Alice"
              data-comment="Fix this."></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);

      // Verify comments.xml exists
      const commentsXml = await zip.file('word/comments.xml').async('string');
      expect(commentsXml).toContain('w:author="Alice"');
      expect(commentsXml).toContain('Fix this.');

      // Verify document.xml has comment range markers
      const docXml = await zip.file('word/document.xml').async('string');
      expect(docXml).toContain('<w:commentRangeStart');
      expect(docXml).toContain('<w:commentRangeEnd');
      expect(docXml).toContain('<w:commentReference');
    });

    test('should not generate comments.xml when no comments present', async () => {
      const html = '<p>Normal text</p>';
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);

      expect(zip.file('word/comments.xml')).toBeNull();
    });

    test('should add CommentReference and CommentText styles', async () => {
      const html = `<p>
        <span class="comment-start" id="0" data-author="A" data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const stylesXml = await zip.file('word/styles.xml').async('string');

      expect(stylesXml).toContain('CommentReference');
      expect(stylesXml).toContain('CommentText');
    });

    test('should add comment relationship to document.xml.rels', async () => {
      const html = `<p>
        <span class="comment-start" id="0" data-author="A" data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const relsXml = await zip.file('word/_rels/document.xml.rels').async('string');

      expect(relsXml).toContain('comments.xml');
      expect(relsXml).toContain('relationships/comments');
    });

    test('should add comment content type override', async () => {
      const html = `<p>
        <span class="comment-start" id="0" data-author="A" data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const contentTypesXml = await zip.file('[Content_Types].xml').async('string');

      expect(contentTypesXml).toContain('comments+xml');
    });

    test('should derive initials from author name', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="Alice Johnson"
              data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const commentsXml = await zip.file('word/comments.xml').async('string');

      expect(commentsXml).toContain('w:initials="AJ"');
    });
  });

  describe('Threaded comments', () => {
    test('should generate commentsExtended.xml for replies', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="Alice"
              data-comment="Fix this."></span>
        text
        <span class="comment-end" id="0"></span>
        <span class="comment-start" id="1"
              data-author="Bob"
              data-comment="I agree."
              data-parent-id="0"></span>
        <span class="comment-end" id="1"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);

      const extXml = await zip.file('word/commentsExtended.xml').async('string');
      expect(extXml).toContain('w15:commentEx');
      expect(extXml).toContain('w15:paraIdParent');
    });

    test('should mark resolved comments with done=1', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="Alice"
              data-comment="Done."
              data-done="1"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);

      const extXml = await zip.file('word/commentsExtended.xml').async('string');
      expect(extXml).toContain('w15:done="1"');
    });
  });

  describe('Combined track changes and comments', () => {
    test('should handle insertion with comment on same range', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="Reviewer"
              data-comment="Good addition."></span>
        <ins data-author="Editor" datetime="2024-01-01T00:00:00Z">new text</ins>
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);
      const zip = await new JSZip().loadAsync(result);

      // Both tracked change and comments present
      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('<w:commentRangeStart');
      expect(zip.file('word/comments.xml')).not.toBeNull();
    });
  });
});
```

### 6.4 Reference Fixture Validation Test

```javascript
describe('Reference DOCX validation', () => {
  test('generated DOCX should match reference structure', async () => {
    const html = `<p>
      The project
      <del data-author="Editor" datetime="2024-06-15T09:00:00Z">will be</del>
      <ins data-author="Editor" datetime="2024-06-15T09:00:00Z">has been</ins>
      <span class="comment-start" id="0"
            data-author="Reviewer"
            data-date="2024-06-15T10:00:00Z"
            data-comment="Consider rephrasing."></span>
      completed on schedule
      <span class="comment-end" id="0"></span>
      and delivered.
    </p>`;

    const result = await HTMLtoDOCX(html, {});
    const zip = await new JSZip().loadAsync(result);

    // Verify all expected files exist
    expect(zip.file('word/document.xml')).not.toBeNull();
    expect(zip.file('word/comments.xml')).not.toBeNull();
    expect(zip.file('word/commentsExtended.xml')).not.toBeNull();
    expect(zip.file('word/settings.xml')).not.toBeNull();
    expect(zip.file('word/styles.xml')).not.toBeNull();
    expect(zip.file('[Content_Types].xml')).not.toBeNull();
    expect(zip.file('word/_rels/document.xml.rels')).not.toBeNull();

    // Verify settings has trackRevisions
    const settings = await zip.file('word/settings.xml').async('string');
    expect(settings).toContain('trackRevisions');

    // Verify content types
    const contentTypes = await zip.file('[Content_Types].xml').async('string');
    expect(contentTypes).toContain('comments+xml');

    // Verify relationships
    const rels = await zip.file('word/_rels/document.xml.rels').async('string');
    expect(rels).toContain('comments.xml');
  });
});
```

### 6.5 Edge Case Tests

```javascript
describe('Edge cases', () => {
  test('should handle empty ins element', async () => {
    const html = '<p>text <ins data-author="Jane"></ins> more</p>';
    const result = await HTMLtoDOCX(html, {});
    // Should not crash
    expect(result).toBeTruthy();
  });

  test('should handle special characters in comment text', async () => {
    const html = `<p>
      <span class="comment-start" id="0"
            data-author="O'Brien"
            data-comment="Use &amp; instead of &lt;and&gt;"></span>
      text
      <span class="comment-end" id="0"></span>
    </p>`;
    const result = await HTMLtoDOCX(html, {});
    const zip = await new JSZip().loadAsync(result);
    const commentsXml = await zip.file('word/comments.xml').async('string');

    // XML special characters should be properly escaped
    expect(commentsXml).not.toContain('&amp;amp;'); // Not double-escaped
  });

  test('should handle multiple comments on different ranges', async () => {
    const html = `
      <p>
        <span class="comment-start" id="0" data-author="A" data-comment="First"></span>
        first range
        <span class="comment-end" id="0"></span>
        middle text
        <span class="comment-start" id="1" data-author="B" data-comment="Second"></span>
        second range
        <span class="comment-end" id="1"></span>
      </p>`;
    const result = await HTMLtoDOCX(html, {});
    const zip = await new JSZip().loadAsync(result);
    const commentsXml = await zip.file('word/comments.xml').async('string');

    expect(commentsXml).toContain('w:id="0"');
    expect(commentsXml).toContain('w:id="1"');
  });

  test('should handle comment spanning multiple paragraphs', async () => {
    const html = `
      <p>
        <span class="comment-start" id="0" data-author="A" data-comment="Spans paras"></span>
        First paragraph
      </p>
      <p>
        Second paragraph
        <span class="comment-end" id="0"></span>
      </p>`;
    const result = await HTMLtoDOCX(html, {});
    const parsed = await parseDOCX(result);

    expect(parsed.xml).toContain('<w:commentRangeStart');
    expect(parsed.xml).toContain('<w:commentRangeEnd');
  });

  test('should handle ins/del without datetime (uses default)', async () => {
    const html = '<p><ins data-author="Jane">text</ins></p>';
    const result = await HTMLtoDOCX(html, {});
    const parsed = await parseDOCX(result);

    // Should have a date attribute (auto-generated)
    expect(parsed.xml).toMatch(/w:date="[^"]+"/);
  });
});
```

---

## 7. Implementation Order

| Phase | Task | Dependencies |
|-------|------|-------------|
| **1** | Add namespaces (`src/namespaces.js`) | None |
| **2** | Add constants (`src/constants.js`) | None |
| **3** | Create `comments.xml` schema (`src/schemas/comments.js`) | Phase 1 |
| **4** | Create `commentsExtended.xml` schema (`src/schemas/comments-extended.js`) | Phase 1 |
| **5** | Update schema exports (`src/schemas/index.js`) | Phases 3-4 |
| **6** | Update `settings.js` for dynamic `trackRevisions` | Phase 1 |
| **7** | Update `styles.js` for comment styles | Phase 1 |
| **8** | Update `document.template.js` for `w14`/`w15` namespaces | Phase 1 |
| **9** | Add state + methods to `docx-document.js` | Phases 1-8 |
| **10** | Add builder functions to `xml-builder.js` | Phase 9 |
| **11** | Update routing in `render-document-file.js` | Phase 10 |
| **12** | Add file generation to `html-to-docx.js` | Phase 9 |
| **13** | Write unit tests | Phases 10-12 |
| **14** | Create reference DOCX fixture | Phase 13 |
| **15** | End-to-end validation | All |

---

## 8. Files Changed Summary

| File | Change Type | Description |
|------|-------------|-------------|
| `src/namespaces.js` | Modified | Add `w14`, `w15`, `comments`, `commentsExtended` URIs |
| `src/constants.js` | Modified | Add `commentType`, `commentsExtendedType`, etc. |
| `src/schemas/comments.js` | **New** | Template for `comments.xml` |
| `src/schemas/comments-extended.js` | **New** | Template for `commentsExtended.xml` |
| `src/schemas/index.js` | Modified | Export new schemas |
| `src/schemas/settings.js` | Modified | Dynamic `trackRevisions` support |
| `src/schemas/styles.js` | Modified | Add `CommentReference` + `CommentText` styles |
| `src/schemas/document.template.js` | Modified | Add `w14`, `w15` namespace declarations |
| `src/docx-document.js` | Modified | New state, methods for comments/revisions/content types/rels |
| `src/helpers/xml-builder.js` | Modified | New `buildTrackedChange`, `buildCommentRangeStart/End`, `buildRunsFromChildren` |
| `src/helpers/render-document-file.js` | Modified | Route `<ins>`/`<del>`/`<span class="comment-*">` to new builders |
| `src/html-to-docx.js` | Modified | Add `comments.xml`, `commentsExtended.xml` to ZIP |
| `tests/tracked-changes.test.js` | **New** | Track changes unit tests |
| `tests/comments.test.js` | **New** | Comments unit tests |
| `tests/fixtures/tracked-changes-and-comments.docx` | **New** | Reference DOCX fixture |

---

## 9. Risk Assessment

| Risk | Mitigation |
|------|------------|
| Breaking existing `<ins>`/`<del>` rendering | Feature gated on `data-author` presence; without it, behavior is identical |
| ID collisions between tracked changes and comments | Separate ID counters for revisions and comments |
| `paraId` collisions in comments | Generate unique 8-char hex IDs via nanoid or deterministic counter |
| XML special characters in comment text | Already handled by xmlbuilder2's automatic escaping |
| Empty `<span>` removal by HTML minifier | `html-minifier-terser` preserves empty elements by default |
| Word compatibility issues | Generate conformant XML matching reference DOCX; test with fixture |
| `commentsExtended.xml` not recognized by older Word | Uses `mc:Ignorable="w15"` — older versions safely ignore it |

---

## 10. Node API & Browser Compatibility

All implementation must work in both environments since `html-to-docx` ships three bundles: ESM (`html-to-docx.esm.js`), UMD (`html-to-docx.umd.js`), and browser IIFE (`html-to-docx.browser.js`).

### Compatibility Constraints

| Concern | Impact | Mitigation |
|---------|--------|------------|
| `nanoid` used for paraId generation | Already bundled for browser; ESM-only import works in all three builds | No change needed — nanoid@3 used, which has CJS fallback |
| `new Date().toISOString()` for default timestamps | Works identically in Node and browser | No change needed |
| `xmlbuilder2` for XML generation | Already bundled for all builds | No change needed |
| `Buffer` usage | Browser build does not have `Buffer` | Do NOT use `Buffer` in any new code. Use plain strings for paraId hex generation (`parseInt`, `toString(16)`, `padStart`) |
| `crypto.randomUUID()` | Not available in all browsers | Use `nanoid` (already a dependency) for ID generation instead |
| Comment/change detection (checking `data-author` attribute) | Runs on VDOM nodes, not platform APIs | Pure JavaScript string/object checks — cross-platform |
| Dynamic XML string templates | No file I/O; in-memory string building with `xmlbuilder2` | Cross-platform by design |

### Specific Implementation Rules

1. **No `Buffer` in new code.** Use `string` operations for hex ID generation:
   ```javascript
   // GOOD: works in browser and Node
   generateParaId() {
     this.lastParaIdCounter += 1;
     return this.lastParaIdCounter.toString(16).toUpperCase().padStart(8, '0');
   }

   // BAD: Buffer not available in browser
   Buffer.from(id, 'utf-8').toString('hex');
   ```

2. **No `fs` or `path` in library code.** File operations are only in `tests/` and the fixture generator script.

3. **No Node.js-only globals** (`process`, `__dirname`, `require`). The codebase already uses ES modules throughout.

4. **Test with the browser build.** After implementation, verify the browser IIFE bundle includes the new code:
   ```bash
   npm run build:browser
   # Verify the built file contains 'commentRangeStart', 'w:ins', 'w:del', etc.
   grep -c 'commentRangeStart' dist/html-to-docx.browser.js
   ```

5. **JSZip API is identical in browser and Node.** All ZIP manipulation uses `JSZip` which is already the bundled dependency for both environments. Adding `comments.xml` and `commentsExtended.xml` to the ZIP uses the same `zip.folder().file()` API.

6. **`html-minifier-terser` is Node-only** but is only used in `index.js` as a preprocessing step, not in the core conversion pipeline. The browser bundle inlines it. The comment span elements (`<span class="comment-start">`) survive minification since they are not HTML comments and `removeEmptyElements` is not enabled.

### Browser Test

Add a test that exercises the browser-compatible code path:

```javascript
test('should work without Buffer (browser compat)', async () => {
  // Temporarily remove Buffer to simulate browser environment
  const originalBuffer = global.Buffer;
  // Note: Can't fully remove Buffer in Node tests, but verify no Buffer usage in new code
  // The real browser test happens via the built IIFE bundle

  const html = `<p>
    <ins data-author="Jane" datetime="2024-01-01T00:00:00Z">inserted</ins>
    <del data-author="Joe" datetime="2024-01-01T00:00:00Z">deleted</del>
    <span class="comment-start" id="0" data-author="A" data-comment="Note"></span>
    text
    <span class="comment-end" id="0"></span>
  </p>`;

  const result = await HTMLtoDOCX(html, {});
  expect(result).toBeTruthy();
  // Verify it's a valid DOCX (ZIP with expected files)
  const zip = await new JSZip().loadAsync(result);
  expect(zip.file('word/document.xml')).not.toBeNull();
  expect(zip.file('word/comments.xml')).not.toBeNull();
});
```

---

## 11. Non-Goals (Out of Scope)

- **Formatting change tracking** (`w:rPrChange`, `w:pPrChange`) — tracks bold/italic changes, too complex for HTML input
- **Move tracking** (`w:moveFrom`/`w:moveTo`) — no standard HTML representation
- **Table cell revision tracking** (`w:cellDel`, `w:cellIns`) — specialized Word feature
- **`people.xml`** — optional in Office 2013+, not needed for core functionality
- **`commentsExtensible.xml`** (Office 2021+, `w16cex`) — too new, limited compatibility
- **Comments containing rich formatting** (images, hyperlinks) — plain text only for initial implementation
- **Comment anchoring to non-text elements** (images, tables) — text ranges only
