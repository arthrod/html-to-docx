/**
 * Script to create a reference DOCX file with tracked changes and comments.
 *
 * This file generates a valid DOCX (Office Open XML) file that contains:
 * - Normal text
 * - An insertion (tracked change)
 * - A deletion (tracked change)
 * - A comment on a text range
 * - A threaded comment (reply to another comment)
 * - A resolved comment (done=1)
 * - Mixed: insertion with overlapping comment
 *
 * The generated file serves as a reference fixture for testing the
 * html-to-docx tracked changes and comments implementation.
 *
 * Usage: node tests/fixtures/create-reference-docx.js
 */

import JSZip from 'jszip';
import { writeFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Namespaces
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  o: 'urn:schemas-microsoft-com:office:office',
  v: 'urn:schemas-microsoft-com:vml',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  ve: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  vt: 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  w10: 'urn:schemas-microsoft-com:office:word',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
};

// ============================================================================
// [Content_Types].xml
// ============================================================================
const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExtended+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

// ============================================================================
// _rels/.rels
// ============================================================================
const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

// ============================================================================
// word/_rels/document.xml.rels
// ============================================================================
const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
</Relationships>`;

// ============================================================================
// word/settings.xml
// ============================================================================
const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${NS.w}" xmlns:o="${NS.o}" xmlns:r="${NS.r}" xmlns:v="${NS.v}">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:trackRevisions/>
  <w:decimalSymbol w:val="."/>
  <w:listSeparator w:val=","/>
</w:settings>`;

// ============================================================================
// word/styles.xml (includes CommentReference and CommentText styles)
// ============================================================================
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}" xmlns:r="${NS.r}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="240" w:lineRule="atLeast"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="normal"/>
  </w:style>
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
</w:styles>`;

// ============================================================================
// word/fontTable.xml
// ============================================================================
const fontTableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="${NS.w}">
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>`;

// ============================================================================
// word/document.xml
// Contains: normal text, tracked insertion, tracked deletion,
//           comment range, and mixed insertion+comment
// ============================================================================
const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS.w}"
            xmlns:r="${NS.r}"
            xmlns:w14="${NS.w14}"
            xmlns:w15="${NS.w15}"
            xmlns:mc="${NS.mc}"
            mc:Ignorable="w14 w15">
  <w:body>

    <!-- Paragraph 1: Normal text -->
    <w:p w14:paraId="00000001" w14:textId="00000001">
      <w:r>
        <w:t>This is a normal paragraph with no tracked changes or comments.</w:t>
      </w:r>
    </w:p>

    <!-- Paragraph 2: Contains a tracked insertion -->
    <w:p w14:paraId="00000002" w14:textId="00000002">
      <w:r>
        <w:t xml:space="preserve">The project </w:t>
      </w:r>
      <w:ins w:id="1" w:author="Jane Doe" w:date="2024-06-15T10:30:00Z">
        <w:r>
          <w:t>has been completed</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t xml:space="preserve"> on schedule.</w:t>
      </w:r>
    </w:p>

    <!-- Paragraph 3: Contains a tracked deletion -->
    <w:p w14:paraId="00000003" w14:textId="00000003">
      <w:r>
        <w:t xml:space="preserve">Some </w:t>
      </w:r>
      <w:del w:id="2" w:author="Joe Smith" w:date="2024-06-15T12:50:00Z">
        <w:r>
          <w:delText xml:space="preserve">unnecessary </w:delText>
        </w:r>
      </w:del>
      <w:r>
        <w:t>text remains here.</w:t>
      </w:r>
    </w:p>

    <!-- Paragraph 4: Contains adjacent deletion and insertion (replacement) -->
    <w:p w14:paraId="00000004" w14:textId="00000004">
      <w:r>
        <w:t xml:space="preserve">We </w:t>
      </w:r>
      <w:del w:id="3" w:author="Editor" w:date="2024-06-15T09:00:00Z">
        <w:r>
          <w:delText>will deliver</w:delText>
        </w:r>
      </w:del>
      <w:ins w:id="4" w:author="Editor" w:date="2024-06-15T09:00:00Z">
        <w:r>
          <w:t>have delivered</w:t>
        </w:r>
      </w:ins>
      <w:r>
        <w:t xml:space="preserve"> the final report.</w:t>
      </w:r>
    </w:p>

    <!-- Paragraph 5: Contains a comment on a text range -->
    <w:p w14:paraId="00000005" w14:textId="00000005">
      <w:r>
        <w:t xml:space="preserve">Please review </w:t>
      </w:r>
      <w:commentRangeStart w:id="10"/>
      <w:r>
        <w:t>this section carefully</w:t>
      </w:r>
      <w:commentRangeEnd w:id="10"/>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:commentReference w:id="10"/>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> before submitting.</w:t>
      </w:r>
    </w:p>

    <!-- Paragraph 6: Contains a comment with a threaded reply -->
    <w:p w14:paraId="00000006" w14:textId="00000006">
      <w:commentRangeStart w:id="11"/>
      <w:r>
        <w:t>The budget figures need updating for Q3.</w:t>
      </w:r>
      <w:commentRangeEnd w:id="11"/>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:commentReference w:id="11"/>
      </w:r>
    </w:p>

    <!-- Paragraph 7: Contains a resolved comment -->
    <w:p w14:paraId="00000007" w14:textId="00000007">
      <w:commentRangeStart w:id="13"/>
      <w:r>
        <w:t>Fixed the typo in the introduction.</w:t>
      </w:r>
      <w:commentRangeEnd w:id="13"/>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:commentReference w:id="13"/>
      </w:r>
    </w:p>

    <!-- Paragraph 8: Mixed - insertion with overlapping comment -->
    <w:p w14:paraId="00000008" w14:textId="00000008">
      <w:r>
        <w:t xml:space="preserve">The deadline is </w:t>
      </w:r>
      <w:commentRangeStart w:id="14"/>
      <w:ins w:id="5" w:author="Manager" w:date="2024-06-16T08:00:00Z">
        <w:r>
          <w:t>next Friday</w:t>
        </w:r>
      </w:ins>
      <w:commentRangeEnd w:id="14"/>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:commentReference w:id="14"/>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>

    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

// ============================================================================
// word/comments.xml
// Contains all comments referenced in document.xml
// ============================================================================
const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${NS.w}"
            xmlns:w14="${NS.w14}"
            xmlns:r="${NS.r}">

  <!-- Comment 10: Basic comment on "this section carefully" -->
  <w:comment w:id="10" w:author="Alice Johnson" w:initials="AJ" w:date="2024-06-15T10:00:00Z">
    <w:p w14:paraId="0000C010" w14:textId="0000T010">
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
        <w:t>This section needs more supporting data. Please add references.</w:t>
      </w:r>
    </w:p>
  </w:comment>

  <!-- Comment 11: Top-level comment on budget figures (has a reply) -->
  <w:comment w:id="11" w:author="Alice Johnson" w:initials="AJ" w:date="2024-06-15T14:00:00Z">
    <w:p w14:paraId="0000C011" w14:textId="0000T011">
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
        <w:t>These numbers are from Q2. Can you update them to Q3?</w:t>
      </w:r>
    </w:p>
  </w:comment>

  <!-- Comment 12: Reply to comment 11 (threaded) -->
  <w:comment w:id="12" w:author="Bob Williams" w:initials="BW" w:date="2024-06-15T15:30:00Z">
    <w:p w14:paraId="0000C012" w14:textId="0000T012">
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
        <w:t>Sure, I will update the figures by end of day.</w:t>
      </w:r>
    </w:p>
  </w:comment>

  <!-- Comment 13: Resolved comment -->
  <w:comment w:id="13" w:author="Carol Davis" w:initials="CD" w:date="2024-06-14T09:00:00Z">
    <w:p w14:paraId="0000C013" w14:textId="0000T013">
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
        <w:t>Good catch, typo has been fixed.</w:t>
      </w:r>
    </w:p>
  </w:comment>

  <!-- Comment 14: Comment on inserted text (mixed scenario) -->
  <w:comment w:id="14" w:author="Alice Johnson" w:initials="AJ" w:date="2024-06-16T08:15:00Z">
    <w:p w14:paraId="0000C014" w14:textId="0000T014">
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
        <w:t>Please confirm this deadline with the client.</w:t>
      </w:r>
    </w:p>
  </w:comment>

</w:comments>`;

// ============================================================================
// word/commentsExtended.xml
// Threading and resolution info for comments
// ============================================================================
const commentsExtendedXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="${NS.w15}"
                xmlns:mc="${NS.mc}"
                mc:Ignorable="w15">
  <!-- Comment 10: top-level, open -->
  <w15:commentEx w15:paraId="0000C010" w15:done="0"/>

  <!-- Comment 11: top-level, open (has reply) -->
  <w15:commentEx w15:paraId="0000C011" w15:done="0"/>

  <!-- Comment 12: reply to comment 11 -->
  <w15:commentEx w15:paraId="0000C012" w15:paraIdParent="0000C011" w15:done="0"/>

  <!-- Comment 13: resolved -->
  <w15:commentEx w15:paraId="0000C013" w15:done="1"/>

  <!-- Comment 14: top-level, open -->
  <w15:commentEx w15:paraId="0000C014" w15:done="0"/>
</w15:commentsEx>`;

// ============================================================================
// docProps/core.xml
// ============================================================================
const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="${NS.cp}"
                   xmlns:dc="${NS.dc}"
                   xmlns:dcterms="${NS.dcterms}"
                   xmlns:dcmitype="${NS.dcmitype}"
                   xmlns:xsi="${NS.xsi}">
  <dc:title>Reference: Tracked Changes and Comments</dc:title>
  <dc:creator>html-to-docx test fixture generator</dc:creator>
  <dc:description>Reference DOCX with tracked changes (insertions, deletions) and threaded comments</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-06-15T08:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-06-16T10:00:00Z</dcterms:modified>
</cp:coreProperties>`;

// ============================================================================
// Build the DOCX ZIP
// ============================================================================
async function buildDocx() {
  const zip = new JSZip();

  zip.file('[Content_Types].xml', contentTypesXml);
  zip.folder('_rels').file('.rels', relsXml);
  zip.folder('word').file('document.xml', documentXml);
  zip.folder('word').file('styles.xml', stylesXml);
  zip.folder('word').file('settings.xml', settingsXml);
  zip.folder('word').file('fontTable.xml', fontTableXml);
  zip.folder('word').file('comments.xml', commentsXml);
  zip.folder('word').file('commentsExtended.xml', commentsExtendedXml);
  zip.folder('word').folder('_rels').file('document.xml.rels', documentRelsXml);
  zip.folder('docProps').file('core.xml', coreXml);

  const buffer = await zip.generateAsync({ type: 'nodebuffer' });
  const outputPath = join(__dirname, 'tracked-changes-and-comments.docx');
  writeFileSync(outputPath, buffer);
  console.log(`Reference DOCX created at: ${outputPath}`);
  console.log(`File size: ${buffer.length} bytes`);
}

buildDocx().catch(console.error);
