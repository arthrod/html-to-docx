import JSZip from 'jszip';
import HTMLtoDOCX from '../index.js';
import { parseDOCX } from './helpers/docx-assertions.js';

describe('Track Changes', () => {
  describe('Insertions', () => {
    test('should generate w:ins for <ins data-author="...">', async () => {
      const html =
        '<p>Hello <ins data-author="Jane" datetime="2024-01-01T00:00:00Z">world</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('w:author="Jane"');
      expect(parsed.xml).toContain('world');
    });

    test('should not generate tracked change for <ins> without data-author', async () => {
      const html = '<p>Hello <ins>world</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      // Without data-author, <ins> renders as plain text (visual-only mode)
      expect(parsed.xml).not.toContain('<w:ins ');
      expect(parsed.xml).toContain('world');
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
      const html =
        '<p><ins data-author="Jane" datetime="2024-01-01T00:00:00Z"><strong>bold inserted</strong></ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('<w:b');
    });

    test('should preserve datetime in w:date attribute', async () => {
      const html =
        '<p><ins data-author="Jane" datetime="2024-06-15T10:30:00Z">text</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('w:date="2024-06-15T10:30:00Z"');
    });
  });

  describe('Deletions', () => {
    test('should generate w:del with w:delText for <del data-author="...">', async () => {
      const html =
        '<p>Some <del data-author="Joe" datetime="2024-01-01T00:00:00Z">removed</del> text</p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:del ');
      expect(parsed.xml).toContain('w:author="Joe"');
      expect(parsed.xml).toContain('<w:delText');
      expect(parsed.xml).toContain('removed');
    });

    test('should not generate tracked change for <del> without data-author', async () => {
      const html = '<p>Some <del>removed</del> text</p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      // Without data-author, <del> renders as plain text (visual-only mode)
      expect(parsed.xml).not.toContain('<w:del ');
      expect(parsed.xml).toContain('removed');
    });

    test('should use w:delText instead of w:t for deleted text', async () => {
      const html =
        '<p><del data-author="Joe" datetime="2024-01-01T00:00:00Z">deleted content</del></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      // Inside a w:del, we should find w:delText
      expect(parsed.xml).toContain('w:delText');
    });
  });

  describe('Settings', () => {
    test('should add trackRevisions to settings.xml when tracked changes present', async () => {
      const html = '<p><ins data-author="Jane">text</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const settingsXml = await zip.file('word/settings.xml').async('string');

      expect(settingsXml).toContain('trackRevisions');
    });

    test('should NOT add trackRevisions when no tracked changes', async () => {
      const html = '<p>Normal text</p>';
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const settingsXml = await zip.file('word/settings.xml').async('string');

      expect(settingsXml).not.toContain('trackRevisions');
    });

    test('should add trackRevisions for deletions too', async () => {
      const html = '<p><del data-author="Joe">text</del></p>';
      const result = await HTMLtoDOCX(html, {});
      const zip = await new JSZip().loadAsync(result);
      const settingsXml = await zip.file('word/settings.xml').async('string');

      expect(settingsXml).toContain('trackRevisions');
    });
  });

  describe('Mixed insertions and deletions', () => {
    test('should handle adjacent ins and del', async () => {
      const html =
        '<p><del data-author="Ed" datetime="2024-01-01T00:00:00Z">old</del><ins data-author="Ed" datetime="2024-01-01T00:00:00Z">new</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:del ');
      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('w:delText');
    });

    test('should handle ins and del with normal text between', async () => {
      const html =
        '<p>Before <del data-author="Ed">old</del> middle <ins data-author="Ed">new</ins> after</p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:del ');
      expect(parsed.xml).toContain('<w:ins ');
      expect(parsed.xml).toContain('Before');
      expect(parsed.xml).toContain('after');
    });
  });

  describe('Edge cases', () => {
    test('should handle empty ins element', async () => {
      const html = '<p>text <ins data-author="Jane"></ins> more</p>';
      const result = await HTMLtoDOCX(html, {});
      expect(result).toBeTruthy();
    });

    test('should handle empty del element', async () => {
      const html = '<p>text <del data-author="Jane"></del> more</p>';
      const result = await HTMLtoDOCX(html, {});
      expect(result).toBeTruthy();
    });

    test('should handle ins/del without datetime (uses default)', async () => {
      const html = '<p><ins data-author="Jane">text</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toMatch(/w:date="[^"]+"/);
    });

    test('should handle special characters in author names', async () => {
      const html =
        '<p><ins data-author="O\'Brien &amp; Sons" datetime="2024-01-01T00:00:00Z">text</ins></p>';
      const result = await HTMLtoDOCX(html, {});
      const parsed = await parseDOCX(result);

      expect(parsed.xml).toContain('<w:ins ');
    });
  });
});
