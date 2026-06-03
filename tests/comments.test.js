import JSZip from 'jszip';
import HTMLtoDOCX from '../index.js';
import { parseDOCX } from './helpers/docx-assertions.js';

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

      const commentsXml = await zip.file('word/comments.xml').async('string');
      expect(commentsXml).toContain('w:author="Alice"');
      expect(commentsXml).toContain('Fix this.');

      const docXml = await zip.file('word/document.xml').async('string');
      expect(docXml).toContain('commentRangeStart');
      expect(docXml).toContain('commentRangeEnd');
      expect(docXml).toContain('commentReference');
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

    test('should use provided initials when specified', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="Alice Johnson"
              data-initials="ALJ"
              data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const commentsXml = await zip.file('word/comments.xml').async('string');

      expect(commentsXml).toContain('w:initials="ALJ"');
    });

    test('should include w14:paraId and w14:textId in comment paragraphs', async () => {
      const html = `<p>
        <span class="comment-start" id="0" data-author="A" data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const commentsXml = await zip.file('word/comments.xml').async('string');

      expect(commentsXml).toContain('w14:paraId');
      expect(commentsXml).toContain('w14:textId');
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

    test('should add commentsExtended content type', async () => {
      const html = `<p>
        <span class="comment-start" id="0" data-author="A" data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const contentTypesXml = await zip.file('[Content_Types].xml').async('string');

      expect(contentTypesXml).toContain('commentsExtended+xml');
    });

    test('should add commentsExtended relationship', async () => {
      const html = `<p>
        <span class="comment-start" id="0" data-author="A" data-comment="Note"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const relsXml = await zip.file('word/_rels/document.xml.rels').async('string');

      expect(relsXml).toContain('commentsExtended.xml');
    });
  });

  describe('Multiple comments', () => {
    test('should handle multiple comments on different ranges', async () => {
      const html = `<p>
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
      expect(commentsXml).toContain('First');
      expect(commentsXml).toContain('Second');
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

      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('commentRangeStart');
      expect(zip.file('word/comments.xml')).not.toBeNull();
    });

    test('should handle deletion with comment on same range', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="Reviewer"
              data-comment="Why remove this?"></span>
        <del data-author="Editor" datetime="2024-01-01T00:00:00Z">old text</del>
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);
      const zip = await new JSZip().loadAsync(result);

      expect(parsed.xml).toContain('<w:del ');
      expect(parsed.xml).toContain('commentRangeStart');
      expect(zip.file('word/comments.xml')).not.toBeNull();
    });
  });

  describe('Reference DOCX validation', () => {
    test('generated DOCX should have all required parts', async () => {
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

      expect(zip.file('word/document.xml')).not.toBeNull();
      expect(zip.file('word/comments.xml')).not.toBeNull();
      expect(zip.file('word/commentsExtended.xml')).not.toBeNull();
      expect(zip.file('word/settings.xml')).not.toBeNull();
      expect(zip.file('word/styles.xml')).not.toBeNull();
      expect(zip.file('[Content_Types].xml')).not.toBeNull();
      expect(zip.file('word/_rels/document.xml.rels')).not.toBeNull();

      const settings = await zip.file('word/settings.xml').async('string');
      expect(settings).toContain('trackRevisions');

      const contentTypes = await zip.file('[Content_Types].xml').async('string');
      expect(contentTypes).toContain('comments+xml');
      expect(contentTypes).toContain('commentsExtended+xml');

      const rels = await zip.file('word/_rels/document.xml.rels').async('string');
      expect(rels).toContain('comments.xml');
      expect(rels).toContain('commentsExtended.xml');
    });
  });

  describe('Edge cases', () => {
    test('should handle special characters in comment text', async () => {
      const html = `<p>
        <span class="comment-start" id="0"
              data-author="O'Brien"
              data-comment="Check the quotes"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const commentsXml = await zip.file('word/comments.xml').async('string');

      // Author with apostrophe should be properly escaped in XML
      expect(commentsXml).toContain("O'Brien");
      expect(commentsXml).toContain('Check the quotes');
    });

    test('should use default author when not provided', async () => {
      const html = `<p>
        <span class="comment-start" id="0" data-comment="Note without author"></span>
        text
        <span class="comment-end" id="0"></span>
      </p>`;
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const commentsXml = await zip.file('word/comments.xml').async('string');

      expect(commentsXml).toContain('w:author="Unknown Author"');
    });
  });
});
