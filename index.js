/* eslint-disable no-useless-escape */
import JSZip from 'jszip';
import { minify } from 'html-minifier-terser';

import createDocumentOptionsAndMergeWithDefaults from './src/utils/options-utils';
import addFilesToContainer from './src/html-to-docx';
import { preprocessResumeCSS } from './src/utils/css-preprocessor';

const minifyHTMLString = async (htmlString) => {
  try {
    if (typeof htmlString === 'string' || htmlString instanceof String) {
      const minifiedHTMLString = await minify(htmlString, {
        collapseWhitespace: true,
        removeComments: true,
      });
      return minifiedHTMLString;
    }

    throw new Error('invalid html string');
  } catch (error) {
    return null;
  }
};

async function generateContainer(
  htmlString,
  headerHTMLString,
  documentOptions = {},
  footerHTMLString
) {
  const zip = new JSZip();

  const normalizedDocumentOptions = createDocumentOptionsAndMergeWithDefaults(documentOptions);

  let contentHTML = htmlString;
  let headerHTML = headerHTMLString;
  let footerHTML = footerHTMLString;
  
  // 新增: CSS预处理 - 将<style>标签中的CSS转换为内联样式
  if (contentHTML && typeof contentHTML === 'string' && normalizedDocumentOptions['preprocessing'] && !normalizedDocumentOptions['preprocessing']['skipCSSPreprocessing']) {
    contentHTML = preprocessResumeCSS(contentHTML);
  }
  if (headerHTML && typeof headerHTML === 'string' && normalizedDocumentOptions['preprocessing'] && !normalizedDocumentOptions['preprocessing']['skipCSSPreprocessing']) {
    headerHTML = preprocessResumeCSS(headerHTML);
  }
  if (footerHTML && typeof footerHTML === 'string' && normalizedDocumentOptions['preprocessing'] && !normalizedDocumentOptions['preprocessing']['skipCSSPreprocessing']) {
    footerHTML = preprocessResumeCSS(footerHTML);
  }
  
  if (htmlString && normalizedDocumentOptions['preprocessing'] && !normalizedDocumentOptions['preprocessing']['skipHTMLMinify']) {
    contentHTML = await minifyHTMLString(contentHTML);
  }
  if (headerHTMLString && normalizedDocumentOptions['preprocessing'] && !normalizedDocumentOptions['preprocessing']['skipHTMLMinify']) {
    headerHTML = await minifyHTMLString(headerHTML);
  }
  if (footerHTMLString && normalizedDocumentOptions['preprocessing'] && !normalizedDocumentOptions['preprocessing']['skipHTMLMinify']) {
    footerHTML = await minifyHTMLString(footerHTML);
  }

  await addFilesToContainer(zip, contentHTML, normalizedDocumentOptions, headerHTML, footerHTML);

  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  if (Object.prototype.hasOwnProperty.call(global, 'Buffer')) {
    return Buffer.from(new Uint8Array(buffer));
  }
  if (Object.prototype.hasOwnProperty.call(global, 'Blob')) {
    // eslint-disable-next-line no-undef
    return new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }
  throw new Error(
    'Add blob support using a polyfill eg https://github.com/bjornstar/blob-polyfill'
  );
}

export default generateContainer;
